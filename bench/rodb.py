"""Read-only SQLite connections from a file path, on any OS.

sqlite3 needs a URI for read-only mode, and a raw path doesn't always make a valid one: Windows paths have
backslashes and a drive letter, and a file name can contain '#' or '%' (e.g. "Model #2.xlsm" -> out/Model #2__…/),
which a URI would read as a fragment or an escape. uri() writes the absolute path with forward slashes and
percent-escapes, the form SQLite documents (file:C:/Documents%20and%20Settings/fred/data.db).
"""
import sqlite3
from pathlib import Path
from urllib.parse import quote


def uri(path) -> str:
    return "file:" + quote(Path(path).resolve().as_posix(), safe="/:") + "?mode=ro"


def connect(path, **kw) -> sqlite3.Connection:
    """Waits up to 30 s (sqlite's default is 5) for a writer to finish, e.g. a workbook's link tables being built."""
    kw.setdefault("timeout", 30)
    return sqlite3.connect(uri(path), uri=True, **kw)


def patched(path, values: dict, **kw) -> sqlite3.Connection:
    """A read-only connection whose `cells` shows other values in place of the saved ones ({(sheet, row, col):
    value}), e.g. the Python overlay's results, so code that reads model.db (dcf.py, valuation.py) works on them
    unchanged. A temporary view named cells shadows the table for this connection only; the file is untouched."""
    db = connect(path, **kw)
    db.executescript("""
        CREATE TEMP TABLE patch(sheet TEXT, row INT, col INT, value, PRIMARY KEY(sheet, row, col));
        CREATE TEMP VIEW cells AS SELECT c.sheet, c.row, c.col, c.addr, c.formula,
            CASE WHEN p.sheet IS NULL THEN c.value ELSE p.value END AS value
            FROM main.cells c LEFT JOIN temp.patch p ON p.sheet = c.sheet AND p.row = c.row AND p.col = c.col;""")
    db.executemany("INSERT INTO temp.patch VALUES (?,?,?,?)", [(*k, v) for k, v in values.items()])
    return db
