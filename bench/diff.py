"""What changed between two versions of a workbook (two model.db files from build_map.py).

Line items are matched by (sheet, label, n-th occurrence of that label), not by row number, so inserting
a row doesn't make everything below it look changed. Cells are then compared column by column.
    uv run python bench/diff.py old/model.db new/model.db
"""
import difflib
import json
import re
import sqlite3
import sys
from collections import Counter, defaultdict

from openpyxl.utils import get_column_letter

LIST_CAP = 150
COORDS = re.compile(r" \(\$?[A-Z]{1,3}\$?\d+(?:\.\.\$?[A-Z]{1,3}\$?\d+)?\)")  # "(K12..HX12)" in patterns
EXT_LINK = re.compile(r"\[\d+\]")  # external-workbook index, e.g. [5]Summary!; Excel renumbers these on save


def _formula_key(f: str | None) -> str:
    return EXT_LINK.sub("[n]", COORDS.sub("", f or ""))


def _same(a, b) -> bool:
    if a == b:
        return True
    try:
        fa, fb = float(a), float(b)
    except (TypeError, ValueError):
        return False
    return abs(fa - fb) <= 1e-9 * max(1.0, abs(fa), abs(fb))


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


DATE = re.compile(r"^\d{4}-\d{2}-\d{2}")


def timeline(db, schema: str, sheet: str) -> dict[int, str]:
    """col -> period date (YYYY-MM-DD) from the sheet's timeline row (layout header_row), dates only."""
    lay = db.execute(f"SELECT layout FROM {schema}.sheets WHERE sheet=?", (sheet,)).fetchone()
    hr = json.loads(lay[0] or "{}").get("header_row") if lay else None
    if not hr:
        return {}
    return {c: str(v)[:10] for c, v in db.execute(f"SELECT col, value FROM {schema}.cells WHERE sheet=? AND row=?",
                                                 (sheet, hr)) if DATE.match(str(v or ""))}


def diff(old_db: str, new_db: str) -> dict:
    db = sqlite3.connect(f"file:{new_db}?mode=ro", uri=True)
    db.execute("ATTACH DATABASE ? AS o", (f"file:{old_db}?mode=ro",))

    old_sheets = dict(db.execute("SELECT sheet, state FROM o.sheets"))
    new_sheets = dict(db.execute("SELECT sheet, state FROM main.sheets"))
    sheets = {"added": [s for s in new_sheets if s not in old_sheets],
              "removed": [s for s in old_sheets if s not in new_sheets],
              "visibility": [{"sheet": s, "old": old_sheets[s], "new": st} for s, st in new_sheets.items()
                             if s in old_sheets and old_sheets[s] != st]}

    # Align each sheet's line items like a text diff: equal runs match; a replaced run of the same length is
    # the same rows relabelled; anything else is added / removed.
    def line_items(src):
        by_sheet = defaultdict(list)
        for sh, r, lab, u, pat in db.execute(f"SELECT sheet, row, label, units, patterns FROM {src}.rows ORDER BY sheet, row"):
            by_sheet[sh].append((r, lab, u, pat))
        return by_sheet
    old_items, new_items = line_items("o"), line_items("main")
    pairs, rows_added, rows_removed, renamed = [], [], [], []
    for sh in new_items:
        o_rows, n_rows = old_items.get(sh, []), new_items[sh]
        if sh not in old_items:
            continue  # whole sheet added; reported above
        sm = difflib.SequenceMatcher(None, [x[1] for x in o_rows], [x[1] for x in n_rows], autojunk=False)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal" or (tag == "replace" and i2 - i1 == j2 - j1):
                for o, n in zip(o_rows[i1:i2], n_rows[j1:j2]):
                    pairs.append((sh, n[1], n[2], o[0], n[0], o[3], n[3]))
                    if tag == "replace":
                        renamed.append({"ref": f"{sh}!r{n[0]}", "old": o[1], "new": n[1]})
            else:
                rows_removed += [{"ref": f"{sh}!r{o[0]}", "label": o[1], "units": o[2]} for o in o_rows[i1:i2]]
                rows_added += [{"ref": f"{sh}!r{n[0]}", "label": n[1], "units": n[2]} for n in n_rows[j1:j2]]
    db.execute("CREATE TEMP TABLE m(sheet, label, units, orow, nrow, op, np)")
    db.executemany("INSERT INTO m VALUES (?,?,?,?,?,?,?)", pairs)
    db.execute("CREATE INDEX temp.ix_mo ON m(sheet, orow)")
    n_matched = len(pairs)
    n_moved = sum(1 for p in pairs if p[3] != p[4])

    formula_rows = []
    for s, lab, nrow, op, np in db.execute("SELECT sheet, label, nrow, op, np FROM m WHERE op IS NOT np"):
        if _formula_key(op) != _formula_key(np):
            formula_rows.append({"ref": f"{s}!r{nrow}", "label": lab, "old": op, "new": np})

    # Period label for a column = the sheet's timeline header cell (from layout detection).
    headers: dict[str, dict[int, str]] = {}
    for s, lay in db.execute("SELECT sheet, layout FROM main.sheets"):
        hr = json.loads(lay or "{}").get("header_row")
        if hr:
            headers[s] = {c: str(v)[:10] for c, v in db.execute(
                "SELECT col, value FROM main.cells WHERE sheet=? AND row=?", (s, hr))}

    # A model rolled forward (next year's version) has the same timeline shifted along: compare each period with
    # the same period, matched by the timeline row's dates, not with the same column. Columns outside the
    # timeline (labels, units, totals) match as they are. Periods only in one version aren't compared.
    rolled, colmap = [], []
    for sh in [x for x in new_sheets if x in old_sheets]:
        ot, nt = timeline(db, "o", sh), timeline(db, "main", sh)
        by_date = {d: c for c, d in nt.items()}
        common = [(c, by_date[d]) for c, d in ot.items() if d in by_date]
        if not common or all(a == b for a, b in common):
            continue
        rolled.append({"sheet": sh, "shift": common[0][1] - common[0][0],
                       "dropped": sorted(d for d in ot.values() if d not in by_date),
                       "added": sorted(d for d in nt.values() if d not in set(ot.values()))})
        tl_new = set(nt)
        for (c,) in db.execute("SELECT DISTINCT col FROM o.cells WHERE sheet=?", (sh,)):
            if c in ot:
                if ot[c] in by_date:
                    colmap.append((sh, c, by_date[ot[c]]))
            elif c not in tl_new:
                colmap.append((sh, c, c))
    db.execute("CREATE TEMP TABLE cm(sheet, ocol, ncol)")
    db.executemany("INSERT INTO cm VALUES (?,?,?)", colmap)
    db.execute("CREATE INDEX temp.ix_cm ON cm(sheet, ocol)")
    rolled_sheets = ",".join("'" + r["sheet"].replace("'", "''") + "'" for r in rolled) or "''"

    renamed_rows = {(r["ref"].split("!r")[0], int(r["ref"].split("!r")[1])) for r in renamed}
    inputs, kind_changes, outputs = [], [], defaultdict(list)
    n_inputs = n_kind = n_outputs = n_blanked = 0
    out_by_sheet = Counter()
    cols = "m.sheet, m.label, m.units, m.nrow, nc.col, nc.addr, oc.value, nc.value, oc.formula, nc.formula"
    changed = "oc.value IS NOT nc.value OR (oc.formula IS NULL) != (nc.formula IS NULL)"
    for s, lab, units, nrow, col, addr, ov, nv, of, nf in db.execute(f"""
            SELECT {cols} FROM m JOIN o.cells oc ON oc.sheet = m.sheet AND oc.row = m.orow
            JOIN main.cells nc ON nc.sheet = m.sheet AND nc.row = m.nrow AND nc.col = oc.col
            WHERE m.sheet NOT IN ({rolled_sheets}) AND ({changed})
            UNION ALL
            SELECT {cols} FROM m JOIN o.cells oc ON oc.sheet = m.sheet AND oc.row = m.orow
            JOIN cm ON cm.sheet = m.sheet AND cm.ocol = oc.col
            JOIN main.cells nc ON nc.sheet = m.sheet AND nc.row = m.nrow AND nc.col = cm.ncol
            WHERE m.sheet IN ({rolled_sheets}) AND ({changed})"""):
        item = {"ref": f"{s}!{addr}", "label": lab, "units": units, "period": headers.get(s, {}).get(col),
                "old": ov, "new": nv}
        if (of is None) != (nf is None):
            n_kind += 1
            if len(kind_changes) < LIST_CAP:
                kind_changes.append({**item, "change": "hard-coded" if nf is None else "now a formula",
                                     "old_formula": of, "new_formula": nf})
        elif _same(ov, nv) or (s, nrow) in renamed_rows and isinstance(nv, str):
            continue  # unchanged, or the label cell of a renamed row (reported under rows_renamed)
        elif nf is None:
            n_inputs += 1
            if len(inputs) < LIST_CAP:
                inputs.append(item)
        else:
            n_outputs += 1
            n_blanked += nv in (None, "") and ov not in (None, "")
            out_by_sheet[s] += 1
            outputs[(s, nrow)].append(item)

    # Key outputs: changed formula rows that nothing else depends on (top of the calculation), then the
    # rest by how many cells moved. Show the first changed cell of each as before -> after.
    has_dependents = {(s, r) for s, r in db.execute("SELECT DISTINCT dst_sheet, dst_row FROM main.edges")}
    upstream = Counter({(s, r): n for s, r, n in db.execute(
        "SELECT src_sheet, src_row, COUNT(*) FROM main.edges GROUP BY src_sheet, src_row")})
    ranked = sorted(outputs, key=lambda k: ((k in has_dependents), -upstream[k], -len(outputs[k])))
    key_outputs = []
    for k in ranked[:40]:
        first = outputs[k][0]
        o, n = _num(first["old"]), _num(first["new"])
        key_outputs.append({**first, "ref": f"{k[0]}!r{k[1]}", "cell": first["ref"], "cells_changed": len(outputs[k]),
                            "top_level": k not in has_dependents,
                            "pct": round(100 * (n - o) / abs(o), 2) if o not in (None, 0) and n is not None else None})

    names_changed = [{"name": n, "old": a, "new": b} for n, a, b in db.execute(
        "SELECT n.name, o.ref, n.ref FROM main.names n JOIN o.names o ON o.name = n.name WHERE o.ref != n.ref")]

    total = (len(sheets["added"]) + len(sheets["removed"]) + len(sheets["visibility"]) + len(rows_added)
             + len(rows_removed) + len(renamed) + len(formula_rows) + n_inputs + n_kind + n_outputs + len(names_changed))
    return {
        "identical_content": total == 0,
        "counts": {"line_items_matched": n_matched, "line_items_moved": n_moved,
                   "line_items_added": len(rows_added), "line_items_removed": len(rows_removed), "line_items_renamed": len(renamed),
                   "formula_rows_changed": len(formula_rows), "inputs_changed": n_inputs,
                   "formula_vs_hardcode": n_kind, "calculated_cells_changed": n_outputs,
                   "calculated_now_blank": n_blanked,
                   "names_changed": len(names_changed)},
        "sheets": sheets,
        "timeline_rolled": rolled,
        "inputs": inputs,
        "key_outputs": key_outputs,
        "calculated_by_sheet": dict(out_by_sheet.most_common()),
        "formula_rows": formula_rows[:LIST_CAP],
        "formula_vs_hardcode": kind_changes,
        "rows_added": rows_added[:LIST_CAP],
        "rows_removed": rows_removed[:LIST_CAP],
        "rows_renamed": renamed[:LIST_CAP],
        "names_changed": names_changed[:LIST_CAP],
    }


if __name__ == "__main__":
    d = diff(sys.argv[1], sys.argv[2])
    print(json.dumps({k: d[k] for k in ("identical_content", "counts", "sheets")}, indent=1))
    for k in ("inputs", "key_outputs", "formula_rows"):
        print(k, json.dumps(d[k][:5], default=str, indent=1))
