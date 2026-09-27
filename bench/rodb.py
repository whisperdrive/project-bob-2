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
    return sqlite3.connect(uri(path), uri=True, **kw)
