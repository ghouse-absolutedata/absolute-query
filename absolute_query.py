"""AbsoluteQuery - spreadsheet-style query, insert and update for any SQL Server table.

Paste rows from Excel/CSV straight into the grid, or load rows from a table,
edit cells, and push changes back. Works on Linux and Windows (tkinter + tksheet).

    pip install tksheet pyodbc            # openpyxl optional, for .xlsx import
    python absolute_query.py

Optional defaults for the first connection: DB_SERVER / DB_UID / DB_PWD from env vars or
a .env file next to the app (or in a parent folder). Otherwise add one in the app.

Reads
  Load     table + optional Where + Top. Paste a spreadsheet column (header, then
           one value per line) into Where to get `Header IN (v1, v2, ...)`.
  SQL      any single SELECT / WITH query, Top caps rows. No database or table
           needs picking first: SQL Server reports where each result column comes
           from, and the result edits that table (three-part names switch the
           database). With joins it edits the picked table if it is in the
           result, else the first table whose key columns are all selected.
           Other tables' and expression columns are read-only, never written.
           Updates need the table key in the result.

  Every read runs READ UNCOMMITTED (the same as WITH (NOLOCK) on every table)
  in its own transaction, which is committed. Rows stream in chunks: the status
  bar shows rows read and time, and Cancel stops a long query.

Writes
  Insert   rows with no DB origin (pasted / typed / imported)
  Update   loaded rows whose cells changed - only changed cells are sent

  Every write is one transaction, all-or-nothing. Rows travel as one JSON
  parameter per batch (OPENJSON), so 20k rows is ~1-2s and one round trip per
  2000 rows. On a failure the bad grid row is located and reported, and
  nothing is committed.

Audit columns (matched case-insensitively, only when present on the table)
  insert   CreatedDate = GETDATE(), CreatedBy = OS login user
  update   ModifiedDate = GETDATE(), ModifiedBy = OS login user

Types, keys and lookups
  Typed values must fit the column type (number, date as YYYY-MM-DD, length,
  decimal digits, bit 0/1, GUID); a wrong one is refused. Pasted values are
  kept but marked red with the reason, and Insert / Update stay disabled until
  they are fixed. Hover a header for type, key and lookup details. Key columns
  have accent-coloured headers, lookup columns blue ones with a dropdown.
  Lookup values come from a lookup file (Settings) or else from a foreign key's
  referenced column (when it has at most 5000 values). A foreign key dropdown
  shows a text column of the referenced table beside the key ("3  ·  ISIN"),
  the cell keeps the key. The column is picked automatically (*Name, then
  *Description, then *Code). Right-click a header -> Lookup… to set any
  column's lookup yourself: a table (database.schema.table allowed), its key
  column and the column shown beside it; or no dropdown; or automatic again.
  Remembered per column (Settings resets all). Header -> View lookup values…
  lists every id with its label, searchable; Use writes the id into the
  selected cells. Typing a label (or part of it) finds
  the entry; Enter stores its key ("valoren" -> 6), in the cell or the fx bar. Lookup file, JSON (a two-column SELECT
  shows its second column the same way):
      {"Sales.dbo.Orders.Status": ["Open", "Shipped", "Closed"],
       "dbo.Orders.Currency": {"sql": "SELECT Code FROM dbo.Currency", "database": "Ref"},
       "CountryCode": {"sql": "SELECT Code FROM dbo.Country"}}
  Keys: database.schema.table.column, schema.table.column, table.column or
  column (any table); the most specific wins. A column with a lookup accepts
  only those values (or blank = NULL).

Dry run
  Runs the same statements inside a transaction and rolls back, so type
  conversions, constraints and triggers are validated, then writes a CSV of
  the planned changes to ./dry_run/. Nothing is committed.

Settings (gear icon in the title bar, saved to ~/.absolutequery.json)
  Review before push   Update shows every changed cell as old -> new, Insert shows
                       the new rows; nothing is sent until confirmed
  Theme                Light / Dark
  Lookup file          allowed values per column, see Types, keys and lookups

Spreadsheet
  Name box + fx bar edit the selected cell; status bar shows Count / Sum / Average.
  Fill handle (drag the square on the selection corner; Ctrl+drag copies),
  Ctrl+D / Ctrl+R fill down / right, header menu sort A-Z / Z-A, filter
  (Ctrl+Shift+L), autofit; drag headers to reorder, borders to resize,
  Ctrl+F find, Ctrl+H replace, Ctrl+Z / Ctrl+Y undo / redo, Ctrl+Arrow jump,
  Ctrl+Shift+Arrow extend, Ctrl+Space / Shift+Space select column / row.
  Filtering only hides rows; hidden rows still count for Insert / Update.

Connections (server icon): several saved servers; passwords in the OS keychain.

Sheets: tabs at the bottom, each with its own connection, database, table / query
  and grid. + or Ctrl+T opens a sheet on the same connection, Ctrl+W closes,
  Ctrl+PgUp / Ctrl+PgDn switch. Open sheets are restored on start (source and
  query, not unsaved grid edits); closing with unsaved edits asks first.

Messages: errors (including trigger errors and server PRINT output) show in the
bottom bar; click it for the full text of recent messages.

Empty cell or NULL = SQL NULL. Identity, computed and timestamp columns are read-only.
"""
from __future__ import annotations

import bisect
import copy
import csv
import datetime as dt
import functools
import getpass
import json
import os
import queue
import re
import sys
import threading
import tkinter as tk
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter import font as tkfont

import pyodbc
from tksheet import Sheet, num2alpha
from tksheet import column_headers as _tks_ch, main_table as _tks_mt, row_index as _tks_ri

import connections as conns
import excel_ops
from icons import ICONS

APP_DIR = Path(__file__).resolve().parent
# the installed launcher runs the app from a cache folder; keep user files in its data folder
DRY_RUN_DIR = Path(os.environ.get("ABSOLUTEQUERY_DATA") or APP_DIR) / "dry_run"
VERSION = os.environ.get("ABSOLUTEQUERY_VERSION", "dev")
BATCH = 2000
SPARE_ROWS = 50  # empty rows kept under the data for new records
LOCK_TIMEOUT_MS = 15000
MAX_CHOICES = 5000  # a lookup with more values gets no dropdown (the server still checks keys)

CREATED_DATE = {"createddate", "created_date", "createdon", "created_on", "createdat", "created_at"}
CREATED_BY = {"createdby", "created_by"}
MODIFIED_DATE = {"modifieddate", "modified_date", "modifiedon", "modified_on", "updateddate",
                 "updated_date", "updatedon", "updated_at", "lastmodified", "lastmodifieddate"}
MODIFIED_BY = {"modifiedby", "modified_by", "updatedby", "updated_by"}

# spreadsheet palettes: grey chrome, one green accent, standard sheet fills
LIGHT = {
    "bg": "#f3f3f3", "surface": "#ffffff", "line": "#d4d4d4", "grid": "#e1e1e1", "field": "#c6c6c6",
    "text": "#1f1f1f", "text2": "#444444", "muted": "#7a7a7a", "disabled": "#a6a6a6",
    "button": "#217346", "button_hover": "#185c37",
    "accent": "#217346", "accent_soft": "#e3efe8", "hover": "#e1e1e1", "header": "#f3f3f3",
    "header_sel": "#d2d2d2", "edited": "#fff2cc", "new": "#e2efda", "old": "#fbe2e0",
    "readonly": "#f2f2f2", "error": "#c42b1c", "link": "#1f5fa8",
}
DARK = {
    "bg": "#202020", "surface": "#2b2b2b", "line": "#404040", "grid": "#3a3a3a", "field": "#555555",
    "text": "#e8e8e8", "text2": "#c4c4c4", "muted": "#9a9a9a", "disabled": "#6e6e6e",
    "button": "#107c41", "button_hover": "#0c6534",
    "accent": "#5cb88a", "accent_soft": "#2f4337", "hover": "#3a3a3a", "header": "#252525",
    "header_sel": "#3c3c3c", "edited": "#4a4020", "new": "#24402d", "old": "#4a2626",
    "readonly": "#333333", "error": "#ff7b6e", "link": "#7fb2ea",
}
C = dict(LIGHT)
if os.name == "nt":
    FONT, CELL, MONO = ("Segoe UI", 9), ("Calibri", 11), ("Consolas", 10)
else:  # Carlito is metric-compatible with Calibri
    FONT, CELL, MONO = ("Noto Sans", 9), ("Carlito", 11), ("DejaVu Sans Mono", 10)
APP_NAME = "AbsoluteQuery"
SETTINGS_FILE = Path.home() / ".absolutequery.json"


def load_settings() -> dict:
    s = {"theme": "Light", "review": True, "lookup_file": "", "ime": False,
         "label_display": "same", "label_format": "{id} · {label}", "mismatch": "both"}
    try:
        s.update(json.loads(SETTINGS_FILE.read_text()))
    except (OSError, ValueError):
        pass
    return s


def save_settings(s: dict) -> None:
    try:
        SETTINGS_FILE.write_text(json.dumps(s))
    except OSError:
        pass


def sheet_colors() -> dict:
    return dict(
        table_bg=C["surface"], table_fg=C["text"], table_grid_fg=C["grid"],
        header_bg=C["header"], header_fg=C["text2"], header_grid_fg=C["line"],
        header_border_fg=C["line"], index_border_fg=C["line"],
        index_bg=C["header"], index_fg=C["text2"], index_grid_fg=C["line"],
        top_left_bg=C["header"], top_left_fg=C["line"], outline_color=C["line"],
        header_selected_cells_bg=C["header_sel"], header_selected_cells_fg=C["accent"],
        index_selected_cells_bg=C["header_sel"], index_selected_cells_fg=C["accent"],
        header_selected_columns_bg=C["accent_soft"], header_selected_columns_fg=C["accent"],
        index_selected_rows_bg=C["accent_soft"], index_selected_rows_fg=C["accent"],
        table_selected_cells_border_fg=C["accent"], table_selected_box_cells_fg=C["accent"],
        table_selected_cells_bg=C["accent_soft"], table_selected_cells_fg=C["text"],
        table_selected_rows_bg=C["accent_soft"], table_selected_rows_fg=C["text"],
        table_selected_columns_bg=C["accent_soft"], table_selected_columns_fg=C["text"],
        table_selected_rows_border_fg=C["accent"], table_selected_columns_border_fg=C["accent"],
        frame_bg=C["surface"], vertical_scroll_background=C["bg"],
        horizontal_scroll_background=C["bg"], vertical_scroll_troughcolor=C["bg"],
        horizontal_scroll_troughcolor=C["bg"], vertical_scroll_bordercolor=C["bg"],
        horizontal_scroll_bordercolor=C["bg"],
    )


# ── env / connection ─────────────────────────────────────────────────────────
def load_env() -> None:
    """Fill os.environ from the nearest .env up the tree (does not override)."""
    for d in [APP_DIR, *APP_DIR.parents]:
        f = d / ".env"
        if f.is_file():
            for line in f.read_text(errors="ignore").splitlines():
                m = re.match(r"\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)", line)
                if m:
                    os.environ.setdefault(m.group(1), m.group(2).strip().strip("'\""))
            return


def odbc_driver() -> str:
    installed = pyodbc.drivers()
    for cand in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server",
                 "ODBC Driver 13 for SQL Server", "SQL Server"):
        if cand in installed:
            return cand
    raise RuntimeError("No SQL Server ODBC driver installed")


def q(name: str) -> str:
    return "[" + name.replace("]", "]]") + "]"


def split_table(name: str) -> tuple[str, str]:
    parts = [p.strip("[] ") for p in name.split(".")]
    if len(parts) == 1:
        return "dbo", parts[0]
    return parts[-2], parts[-1]


def fmt(v) -> str:
    """DB value -> grid text. Round-trips through implicit nvarchar conversion."""
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        s = v.strftime("%Y-%m-%d %H:%M:%S")
        return s + f".{v.microsecond // 1000:03d}" if v.microsecond else s
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, Decimal):
        return format(v, "f")
    if isinstance(v, (bytes, bytearray)):
        return "0x" + v.hex()
    return str(v)


class Col:
    def __init__(self, name, type_, is_identity, is_computed, is_pk, external=False,
                 max_length=0, precision=0, scale=0):
        self.name, self.type = name, type_
        # characters for (n)(var)char (-1 = max), digits for decimal / numeric
        self.size = max_length // 2 if type_ in ("nchar", "nvarchar") and max_length > 0 else max_length
        self.precision, self.scale = precision, scale
        self.fk = ""  # "schema.table.column" this column references, if a single-column foreign key
        self.lookup = ""  # where the allowed values come from (foreign key or lookup file)
        self.choices: frozenset[str] | None = None  # allowed values; None = anything of the type
        self.choice_list: list[str] = []  # dropdown entries: "key" or "key  ·  label"
        self.picked: dict[str, str] = {}  # dropdown entry -> key, for entries with a label
        self.key_label: dict[str, str] = {}  # key -> its label fields joined: "Bond · Fixed income"
        self.label_fields: list[str] = []  # names of the fields shown beside the key
        self.key_fields: dict[str, tuple[str, ...]] = {}  # key -> each label field's value
        # table lookup in use (foreign key or set from the header menu):
        # {"db", "table" (schema.table), "key", "label" (field, list of fields, "" = key only)}
        self.src: dict | None = None
        self.of: Col | None = None  # set on a label column: the lookup column it describes
        self.part: int | None = None  # label column of one field: its index in of.label_fields
        self.scratch = False  # a column of your own on this sheet: editable, never written
        # a lookup too big to list (over MAX_CHOICES rows): the server is asked per value.
        # key_fields / key_label then hold the values found so far, `missing` those not in it,
        # `nomatch` typed text that named no single row (-> why)
        self.big = False
        self.key_type = ""  # the lookup key's SQL type, for TRY_CONVERT
        self.missing: set[str] = set()
        self.nomatch: dict[str, str] = {}
        self.search_fields: list[str] = []
        self.subs: dict[int, Col] = {}  # label field index -> its own lookup ("Col ▸ Field")
        self.fmt = ""  # shown text of a sub lookup's value, "" = the raw value (set by App)

        self.label = name  # grid header; a query alias when it differs from the column name
        self.external = external  # query column not in the target table: shown, never written
        self.readonly = bool(external or is_identity or is_computed
                             or type_ in ("timestamp", "rowversion"))
        self.is_pk = bool(is_pk)
        low = name.lower()
        self.audit = ("cdate" if low in CREATED_DATE else "cby" if low in CREATED_BY else
                      "mdate" if low in MODIFIED_DATE else "mby" if low in MODIFIED_BY else None)

    def in_lookup(self, v: str) -> bool | None:
        """Is `v` an allowed value? None: not known (no lookup, or not asked yet)."""
        if self.big:
            return True if v in self.key_fields else False if v in self.missing else None
        return None if self.choices is None else v in self.choices

    @property
    def labelled(self) -> bool:
        """Has label fields to show beside its values."""
        return bool(self.key_label) or (self.big and bool(self.label_fields))


class RowError(Exception):
    def __init__(self, grid_row: int | None, msg: str):
        super().__init__(msg)
        self.grid_row, self.msg = grid_row, msg


class Db:
    def __init__(self):
        self.cn: pyodbc.Connection | None = None
        self.args = None
        self.active: pyodbc.Cursor | None = None  # running statement, for cancel()
        self.fetched = 0  # rows read so far by the running statement, for progress
        self.notes: list[str] = []  # server messages of the last read / write (PRINT, triggers)
        self.cancelled = False

    def connect(self, server, uid, pwd):
        self.close()
        drv = odbc_driver()
        extra = "TrustServerCertificate=yes;Encrypt=no;" if drv.endswith(("18 for SQL Server",)) else ""
        auth = f"UID={uid};PWD={pwd};" if uid else "Trusted_Connection=yes;"
        cs = f"DRIVER={{{drv}}};SERVER={server};DATABASE=master;{auth}{extra}APP=AbsoluteQuery;"
        self.cn = pyodbc.connect(cs, timeout=10, autocommit=False)
        self.cn.timeout = 300
        self.cn.execute(f"SET LOCK_TIMEOUT {LOCK_TIMEOUT_MS}")
        self.cn.commit()
        self.args = (server, uid, pwd)

    def close(self):
        if self.cn is not None:
            try:
                self.cn.close()
            except pyodbc.Error:
                pass
            self.cn = None

    def cur(self, database: str, write: bool = False) -> pyodbc.Cursor:
        """Cursor in `database`. Reads run READ UNCOMMITTED - the same as NOLOCK on
        every table - so they never wait on or block other sessions; writes run
        READ COMMITTED."""
        if self.cancelled:  # the job was cancelled: run nothing more of it
            raise RuntimeError("Cancelled")
        if self.cn is None:
            raise RuntimeError("Not connected")
        setup = ((f"USE {q(database)}; " if database else "") + "SET TRANSACTION ISOLATION LEVEL "
                 f"{'READ COMMITTED' if write else 'READ UNCOMMITTED'};")
        try:
            c = self.active = self.cn.cursor()  # any statement can be cancelled
            c.execute(setup)
            return c
        except pyodbc.Error as e:
            if self.cancelled:
                raise RuntimeError("Cancelled") from None
            if "08S01" in str(e) or "08003" in str(e):  # dropped link: reconnect once
                self.connect(*self.args)
                c = self.active = self.cn.cursor()
                c.execute(setup)
                return c
            raise

    def _read(self, database, sql, params=(), top=None) -> tuple[list[str], list[list[str]]]:
        """One read in its own transaction, committed. Rows arrive in chunks so
        self.fetched shows progress and cancel() can stop it between chunks."""
        c = self.cur(database)
        self.active, self.fetched, self.notes = c, 0, []
        try:
            c.execute(sql, *params)
            self.notes += server_notes(c)
            while c.description is None and c.nextset():
                pass
            if c.description is None:
                raise RuntimeError("Query returned no result set")
            names = [d[0] or f"col{i + 1}" for i, d in enumerate(c.description)]
            rows: list[list[str]] = []
            while top is None or len(rows) < top:
                if self.cancelled:
                    raise RuntimeError("Cancelled")
                chunk = c.fetchmany(5000 if top is None else min(5000, top - len(rows)))
                if not chunk:
                    break
                rows += [[v if type(v) is str else fmt(v) for v in r] for r in chunk]  # most cells are text
                self.fetched = len(rows)
            c.close()
            self.cn.commit()
            return names, rows
        except Exception:
            try:
                self.cn.rollback()
            except pyodbc.Error:
                pass
            if self.cancelled:  # the driver reports a cancel in several ways
                raise RuntimeError("Cancelled") from None
            raise
        finally:
            self.active = None

    def eval_fx(self, database: str, sql: str, items: list[dict]) -> dict[int, str | Exception]:
        """Run a formula query (fx_sql) for `items`, 2,000 cells a query. A failing
        query is split until the failing cells are found; an error in the formula
        itself (syntax, unknown function) is found with one cell and given to all."""
        out: dict[int, str | Exception] = {}
        budget = [200]  # queries spent on isolating errors

        def run(chunk):
            try:
                for r in self._read(database, sql, (json.dumps(chunk),))[1]:
                    out[int(r[0])] = r[1]
                return
            except pyodbc.Error as e:
                if self.cancelled:
                    raise
                err = e
            if len(chunk) == 1 or budget[0] <= 0:
                out.update({it["i"]: err for it in chunk})
                return
            budget[0] -= 1
            if "42000" in str(err) or "42S" in str(err):  # compile error: every cell has it
                out.update({it["i"]: err for it in chunk})
                return
            half = len(chunk) // 2
            run(chunk[:half])
            run(chunk[half:])

        for k in range(0, len(items), 2000):
            run(items[k:k + 2000])
        return out

    def begin(self):
        """A new job starts: earlier cancels no longer apply."""
        self.cancelled = False

    def cancel(self):
        """Stop the running job (UI thread): its current statement gets a
        cancel request and no further statement of it runs. The request goes
        from a helper thread, as a dead network can block it too."""
        self.cancelled = True
        c = self.active
        if c is not None:
            def send():
                try:
                    c.cancel()
                except pyodbc.Error:
                    pass
            threading.Thread(target=send, daemon=True).start()

    def databases(self) -> list[str]:
        c = self.cur("master")
        rows = [r[0] for r in c.execute(
            "SELECT name FROM sys.databases WHERE state = 0 AND HAS_DBACCESS(name) = 1 "
            "AND database_id > 4 ORDER BY name")]
        self.cn.commit()
        return rows

    def tables(self, database: str) -> list[str]:
        c = self.cur(database)
        rows = [r[0] for r in c.execute(
            "SELECT s.name + '.' + t.name FROM sys.tables t JOIN sys.schemas s "
            "ON s.schema_id = t.schema_id ORDER BY s.name, t.name")]
        self.cn.commit()
        return rows

    def columns(self, database: str, table: str) -> list[Col]:
        schema, name = split_table(table)
        c = self.cur(database)
        rows = c.execute("""
            SELECT c.name, ty.name, c.is_identity, c.is_computed,
                   CASE WHEN pk.column_id IS NULL THEN 0 ELSE 1 END,
                   c.max_length, c.precision, c.scale
            FROM sys.columns c
            JOIN sys.types ty ON ty.user_type_id = c.user_type_id
            LEFT JOIN (SELECT ic.object_id, ic.column_id
                       FROM sys.indexes i
                       JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id
                       WHERE i.is_primary_key = 1) pk
                   ON pk.object_id = c.object_id AND pk.column_id = c.column_id
            WHERE c.object_id = OBJECT_ID(?)
            ORDER BY c.column_id""", f"{q(schema)}.{q(name)}").fetchall()
        if rows and not any(r[4] for r in rows):  # no PK: fall back to first unique index
            uq = c.execute("""
                SELECT COL_NAME(ic.object_id, ic.column_id) FROM sys.indexes i
                JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id
                WHERE i.object_id = OBJECT_ID(?) AND i.is_unique = 1 AND ic.is_included_column = 0
                  AND i.index_id = (SELECT MIN(index_id) FROM sys.indexes
                                    WHERE object_id = OBJECT_ID(?) AND is_unique = 1)""",
                f"{q(schema)}.{q(name)}", f"{q(schema)}.{q(name)}").fetchall()
            uqs = {r[0] for r in uq}
            rows = [(r[0], r[1], r[2], r[3], r[0] in uqs, *r[5:]) for r in rows]
        self.cn.commit()
        if not rows:
            raise RuntimeError(f"Table {database}.{schema}.{name} not found")
        cols = [Col(*r[:5], max_length=r[5], precision=r[6], scale=r[7]) for r in rows]
        fks = c.execute("""
            SELECT COL_NAME(k.parent_object_id, k.parent_column_id),
                   OBJECT_SCHEMA_NAME(k.referenced_object_id) + '.' + OBJECT_NAME(k.referenced_object_id)
                   + '.' + COL_NAME(k.referenced_object_id, k.referenced_column_id)
            FROM sys.foreign_key_columns k
            WHERE k.parent_object_id = OBJECT_ID(?)
              AND NOT EXISTS (SELECT 1 FROM sys.foreign_key_columns o  -- single-column keys only
                              WHERE o.constraint_object_id = k.constraint_object_id
                                AND o.constraint_column_id > 1)""", f"{q(schema)}.{q(name)}").fetchall()
        self.cn.commit()
        by_name = {col.name: col for col in cols}
        for col_name, ref in fks:
            if col_name in by_name:
                by_name[col_name].fk = ref
        return cols

    def lookups(self, database: str, table: str, cols: list[Col], spec: dict, prefs: dict) -> None:
        """Fill each column's allowed values, first match wins:
        1. the lookup set from the header menu (prefs["lookups"], see col_key),
        2. the lookup file (`spec`, see lookup_for),
        3. the foreign key's referenced table.
        A table lookup shows a text column beside the key: the one chosen, else
        the likeliest (guess_label)."""
        overrides, fk_labels = prefs.get("lookups", {}), prefs.get("fk_labels", {})
        jobs, tables = [], []
        for col in cols:
            col.choices, col.choice_list, col.picked, col.key_label, col.src = None, [], {}, {}, None
            col.label_fields, col.key_fields = [], {}
            col.big, col.key_type, col.missing, col.nomatch = False, "", set(), {}
            col.lookup = ""
            if col.external or col.readonly:
                continue
            over = overrides.get(col_key(database, table, col.name))
            if over:
                if not over.get("off"):
                    tables.append((col, {k: over.get(k) for k in ("db", "table", "key", "label")},
                                   (f"{over['db']}." if (over.get("db") or database).lower() != database.lower()
                                    else "") + f"{over['table']}.{over['key']} (set by you)"))
                continue
            entry = lookup_for(spec, database, table, col.name)
            if isinstance(entry, list):
                set_choices(col, [str(v) for v in entry], "lookup file")
            elif isinstance(entry, dict) and entry.get("sql"):
                jobs.append((col, entry.get("database") or database, entry["sql"], "lookup file"))
            elif col.fk:
                ref_schema, ref_table, ref_col = col.fk.split(".")
                label = fk_labels.get(f"{database}.{ref_schema}.{ref_table}")
                if label != NO_LOOKUP:
                    tables.append((col, {"db": database, "table": f"{ref_schema}.{ref_table}", "key": ref_col,
                                         "label": label}, f"foreign key to {col.fk}"))
        for col, db, sql, source in jobs:
            err = check_select(sql)
            if err:
                raise RuntimeError(f"Lookup for {col.name}: {err}")
            set_choices(col, self._read(db, sql, top=MAX_CHOICES + 1)[1], source)
        for col, src, source in tables:
            self.table_lookup(col, src, source)
        for col in cols:
            self.sub_lookups(col, col.name, overrides, database, table)

    def sub_lookups(self, col: Col, name: str, overrides: dict, database: str, table: str, depth: int = 0):
        """Lookups set on label fields ("IssuerID ▸ SectorID"), and on theirs in turn.
        A small table is read once, a big one per value in use (App._big_lookups)."""
        col.subs = {}
        if depth > 4:
            return
        for part, field in enumerate(col.label_fields):
            over = overrides.get(col_key(database, table, f"{name} ▸ {field}"))
            if not over or over.get("off"):
                continue
            sub = Col(f"{name} ▸ {field}", "", False, False, False, external=True)
            self.table_lookup(sub, {k: over.get(k) for k in ("db", "table", "key", "label")},
                              f"{over['db']}.{over['table']}.{over['key']} (set by you)")
            col.subs[part] = sub
            self.sub_lookups(sub, sub.name, overrides, database, table, depth + 1)

    def table_lookup(self, col: Col, src: dict, source: str) -> None:
        """Allowed values of `col` from a lookup table; none when it has more
        rows than a dropdown can list (the server still checks a foreign key)."""
        col.src = dict(src)
        schema, name = split_table(src["table"])
        obj = f"{q(schema)}.{q(name)}"
        rows = self._read(src["db"], "SELECT SUM(rows) FROM sys.partitions WHERE object_id = OBJECT_ID(?) "
                                     "AND index_id IN (0, 1)", (obj,))[1]
        if not rows or rows[0][0] is None:
            raise RuntimeError(f"Lookup table {src['db']}.{src['table']} not found")
        if col.src.get("label") is None:
            col.src["label"] = guess_label(self.text_columns(src["db"], src["table"], src["key"]), src["table"])
        fields = label_fields(col.src)
        if int(rows[0][0]) > MAX_CHOICES:  # too many to list: labels are read per value
            t = self._read(src["db"], "SELECT TYPE_NAME(user_type_id), max_length, precision, scale "
                                      "FROM sys.columns WHERE object_id = OBJECT_ID(?) AND name = ?",
                           (obj, src["key"]))[1]
            if not t:
                raise RuntimeError(f"Lookup key {src['key']} not found in {src['db']}.{src['table']}")
            kt = Col("k", t[0][0], False, False, False, max_length=int(t[0][1]), precision=int(t[0][2]),
                     scale=int(t[0][3]))
            col.big, col.label_fields, col.lookup = True, fields, f"{source}, {int(rows[0][0]):,} rows"
            texts = set(self.text_columns(src["db"], src["table"]))
            col.search_fields = [f for f in fields if f in texts]  # names to find a typed value by
            col.key_type = (f"{kt.type}({'max' if kt.size == -1 else kt.size})" if kt.type in TEXT_TYPES
                            else sql_type(kt) or "nvarchar(4000)")
            return
        set_choices(col, self._read(src["db"], lookup_sql(src["table"], src["key"], fields),
                                    top=MAX_CHOICES + 1)[1], source, fields)

    def big_labels(self, col: Col, values: list[str]) -> tuple[dict[str, tuple], set[str]]:
        """Rows of a big lookup for grid `values`: ({value: label fields}, values not in it).
        One query per 2,000 values; the key is compared in its own type (TRY_CONVERT), so an
        index on it is used and text that is no key just finds nothing."""
        src, fields = col.src, col.label_fields
        schema, name = split_table(src["table"])
        cols = ", ".join(q(c) for c in (src["key"], *fields))
        tbl = f"{q(schema)}.{q(name)}"
        sql = (f"SELECT {', '.join('t.' + q(c) for c in (src['key'], *fields))} FROM "
               f"(SELECT DISTINCT TRY_CONVERT({col.key_type}, value) AS k FROM OPENJSON(?)) j "
               f"JOIN {tbl} t WITH (FORCESEEK) ON t.{q(src['key'])} = j.k")
        # a seek per value: a cold plan otherwise may scan the whole table; FORCESEEK
        # needs an index on the key, so without one the plain join is used
        found: dict[str, tuple] = {}
        by_norm: dict[str, list[str]] = {}
        for v in values:
            by_norm.setdefault(v.strip().lower(), []).append(v)
        for i in range(0, len(values), 2000):
            chunk = (json.dumps(values[i:i + 2000]),)
            try:
                rows = self._read(src["db"], sql, chunk)[1]
            except pyodbc.Error as e:
                if "8622" not in str(e) or self.cancelled:  # 8622: no index to seek
                    raise
                sql = sql.replace(" WITH (FORCESEEK)", "")
                rows = self._read(src["db"], sql, chunk)[1]
            for r in rows:
                for v in by_norm.get(str(r[0]).strip().lower(), []):  # same key as the server sees it
                    found.setdefault(v, tuple(r[1:]))
        return found, {v for v in values if v not in found}

    def big_search(self, col: Col, text: str, top: int = 200, exact: bool = False) -> list[list[str]]:
        """Rows of a big lookup whose key or a label field matches `text`: equal to it
        (`exact`, case as the server compares) or containing it."""
        src, fields = col.src, col.label_fields
        schema, name = split_table(src["table"])
        cols = ", ".join(q(c) for c in (src["key"], *fields))
        if not text:
            return self._read(src["db"], f"SELECT TOP ({top}) {cols} FROM {q(schema)}.{q(name)} "
                                         f"ORDER BY {q(src['key'])}", top=top)[1]
        like = text if exact else "%" + re.sub(r"([%_\[])", r"[\1]", text) + "%"
        op = "=" if exact else "LIKE"
        tests = [f"{q(c)} {op} ?" for c in col.search_fields]
        if not tests or re.fullmatch(r"-?\d+", text) or col.key_type.startswith(("varchar", "nvarchar", "char")):
            tests.append(f"CONVERT(nvarchar(4000), {q(src['key'])}) {op} ?")  # key typed: numbers or text keys
        return self._read(src["db"], f"SELECT TOP ({top}) {cols} FROM {q(schema)}.{q(name)} "
                                     f"WHERE {' OR '.join(tests)} ORDER BY {q(src['key'])}",
                          [like] * len(tests), top=top)[1]

    def big_resolve(self, col: Col, text: str) -> tuple[str, tuple] | str:
        """Typed text -> (key, label fields) of the one row it names, else why not.
        A text field equal to the text wins; else the one row starting with it; else
        the one row containing it. Each step is one query (the last scans the table)."""
        src, fields = col.src, col.label_fields
        if not col.search_fields:
            return f"'{text}' is not a key of {src_name(col)}"
        schema, name = split_table(src["table"])
        cols = ", ".join(q(c) for c in (src["key"], *fields))
        pat = re.sub(r"([%_\[])", r"[\1]", text)
        exact = " + ".join(f"CASE WHEN {q(f)} = ? THEN 1 ELSE 0 END" for f in col.search_fields)
        for like in (pat + "%", "%" + pat + "%"):
            tests = " OR ".join(f"{q(f)} LIKE ?" for f in col.search_fields)
            rows = self._read(src["db"], f"SELECT TOP (6) {cols}, {exact} FROM {q(schema)}.{q(name)} "
                                         f"WHERE {tests} ORDER BY {exact} DESC",
                              [text] * len(col.search_fields) + [like] * len(col.search_fields)
                              + [text] * len(col.search_fields), top=6)[1]
            hits = [r for r in rows if r[-1] not in ("0", 0)] or rows  # equal beats starting / containing
            keys = {str(r[0]): r[:-1] for r in hits}
            if len(keys) == 1:
                k, r = next(iter(keys.items()))
                return k, tuple(r[1:])
            if len(keys) > 1:
                names = ", ".join(f"{k} {LABEL_SEP.join(str(x) for x in r[1:] if x)}".strip()
                                  for k, r in list(keys.items())[:3])
                more = "+" if len(keys) > 5 else ""
                return f"'{text}' matches {len(keys)}{more} rows ({names}…): use View lookup values"
        return f"'{text}' is not in {src_name(col)}"

    def text_columns(self, database: str, table: str, but: str = "") -> list[str]:
        schema, name = split_table(table)
        return [r[0] for r in self._read(database, """
            SELECT c.name FROM sys.columns c
            JOIN sys.types t ON t.user_type_id = c.user_type_id
            WHERE c.object_id = OBJECT_ID(?) AND c.name <> ?
              AND t.name IN ('char', 'varchar', 'nchar', 'nvarchar')
            ORDER BY c.column_id""", (f"{q(schema)}.{q(name)}", but))[1]]

    def load(self, database, table, cols: list[Col], where: str, top: int):
        schema, name = split_table(table)
        sel = ", ".join(q(c.name) for c in cols)
        sql = f"SELECT TOP ({int(top)}) {sel} FROM {q(schema)}.{q(name)}"
        if where.strip():
            sql += f" WHERE {where}"
        keys = [q(c.name) for c in cols if c.is_pk]
        if keys:
            sql += " ORDER BY " + ", ".join(keys)
        return self._read(database, sql, top=top)[1]

    def sources(self, database, sql: str) -> list[tuple] | None:
        """Real origin of each result column: (database, schema, table, column) or
        None for expressions. Uses SQL Server's own metadata, so aliases, joins and
        three-part names resolve exactly. None when the server cannot describe it."""
        c = self.cur(database)
        try:
            rows = c.execute(
                "SELECT source_database, source_schema, source_table, source_column "
                "FROM sys.dm_exec_describe_first_result_set(?, NULL, 1) "
                "WHERE is_hidden = 0 ORDER BY column_ordinal", sql).fetchall()
        except pyodbc.Error:
            rows = None
        finally:
            c.close()
            self.cn.commit()
        if not rows:
            return None
        if not database:
            database = self.cn.execute("SELECT DB_NAME()").fetchval()
            self.cn.commit()
        return [(r[0] or database, r[1] or "dbo", r[2], r[3]) if r[2] and r[3] else None for r in rows]

    def query(self, database, sql: str, top: int) -> tuple[list[str], list[list[str]]]:
        """Run a user SELECT (checked read-only by check_select), first `top` rows."""
        return self._read(database, sql, top=top)

    # writes ---------------------------------------------------------------
    def write(self, database, steps, user: str, dry_run=False) -> int:
        """Run every step in ONE transaction, commit at the end (rollback on dry run).
        step = (sql, payloads, expect_exact); sql takes params (user, json batch),
        payloads = [(grid_row, json_obj)]. Returns affected rows."""
        c = self.cur(database, write=True)
        total = 0
        self.notes = []
        try:
            for sql, payloads, exact in steps:
                for i in range(0, len(payloads), BATCH):
                    chunk = payloads[i:i + BATCH]
                    try:
                        c.execute(sql, user, json.dumps([p for _, p in chunk], ensure_ascii=False))
                    except pyodbc.Error as e:
                        self.cn.rollback()
                        raise self._locate(database, sql, chunk, user, e)
                    n = c.rowcount
                    self.notes += server_notes(c)
                    if exact and n != len(chunk):
                        raise RowError(None, f"Expected {len(chunk)} rows affected, got {n}. "
                                             "Key is not unique or rows changed on the server.")
                    total += n if n >= 0 else len(chunk)
            if dry_run:
                self.cn.rollback()
            else:
                self.cn.commit()
        except Exception:
            try:
                self.cn.rollback()
            except pyodbc.Error:
                pass
            raise
        return total

    def _locate(self, database, sql, chunk, user, err) -> RowError:
        """Replay a failed batch row by row inside a rolled-back txn to find the bad row."""
        c = self.cur(database, write=True)
        try:
            for grid_row, p in chunk:
                try:
                    c.execute(sql, user, json.dumps([p], ensure_ascii=False))
                except pyodbc.Error as e:
                    return RowError(grid_row, clean_err(e))
        finally:
            self.cn.rollback()
        return RowError(None, clean_err(err))


def server_notes(c: pyodbc.Cursor) -> list[str]:
    """Informational messages the statement produced (PRINT, RAISERROR below
    severity 11, trigger output), cleaned like errors."""
    return [clean_err(RuntimeError(m[1])) for m in (getattr(c, "messages", None) or [])]


def clean_err(e: Exception) -> str:
    s = str(e.args[1] if isinstance(e, pyodbc.Error) and len(e.args) > 1 else e)
    if "HY008" in s or "Operation canceled" in s:
        return "Cancelled"
    s = re.sub(r"^\[\w+\]\s*", "", s)
    s = re.sub(r"(\[Microsoft\])?\[ODBC Driver \d+ for SQL Server\](\[SQL Server\])?", "", s)
    s = re.sub(r"\s*\(\d+\)(\s*\(SQL\w+\))?", "\n", s)  # one line per diagnostic record
    lines = [re.sub(r"^;?\s*(\[\w+\]\s*)?", "", ln.strip()) for ln in s.splitlines()]
    lines = [ln for ln in lines if ln]
    return "\n".join(dict.fromkeys(lines))


# ── SQL builders ─────────────────────────────────────────────────────────────
def with_clause(n: int, prefix: str) -> str:
    return ", ".join(f"[{prefix}{i}] nvarchar(max) '$.{prefix}{i}'" for i in range(n))


# Every statement takes params (user, json): @u feeds CreatedBy / ModifiedBy.
DECL = "DECLARE @u nvarchar(128) = ?; "


def build_insert(table, cols: list[Col], table_cols: list[Col]) -> tuple[str, list[Col]]:
    """cols = grid columns, table_cols = all table columns (audit fields are set even
    when the grid does not show them). JSON object keys v0..vn follow data_cols."""
    schema, name = split_table(table)
    data_cols, seen = [], set()
    for c in cols:  # a column shown twice is inserted once
        if not c.readonly and not c.scratch and c.audit is None and c.name.lower() not in seen:
            seen.add(c.name.lower())
            data_cols.append(c)
    tgt = [q(c.name) for c in data_cols]
    src = [f"j.[v{i}]" for i in range(len(data_cols))]
    for c in table_cols:
        if c.readonly:
            continue
        if c.audit == "cdate":
            tgt.append(q(c.name)); src.append("GETDATE()")
        elif c.audit == "cby":
            tgt.append(q(c.name)); src.append("@u")
    sql = (DECL + f"INSERT INTO {q(schema)}.{q(name)} ({', '.join(tgt)}) "
           f"SELECT {', '.join(src)} FROM OPENJSON(?) WITH ({with_clause(len(data_cols), 'v')}) j")
    return sql, data_cols


def build_update(table, table_cols: list[Col], keys: list[Col], changed: list[Col]) -> str:
    schema, name = split_table(table)
    sets = [f"t.{q(c.name)} = j.[v{i}]" for i, c in enumerate(changed)]
    changed_names = {c.name for c in changed}
    for c in table_cols:
        if c.readonly or c.name in changed_names:
            continue
        if c.audit == "mdate":
            sets.append(f"t.{q(c.name)} = GETDATE()")
        elif c.audit == "mby":
            sets.append(f"t.{q(c.name)} = @u")
    on = " AND ".join(f"t.{q(k.name)} = j.[k{i}]" for i, k in enumerate(keys))
    return (DECL + f"UPDATE t SET {', '.join(sets)} FROM {q(schema)}.{q(name)} t "
           f"JOIN OPENJSON(?) WITH ({with_clause(len(keys), 'k')}, {with_clause(len(changed), 'v')}) j "
           f"ON {on}")


def dedupe_edits(edited, rows, orig, kpos, names):
    """A join can show one table row in several grid rows. Per (key, column), the
    same new value is sent once; different new values are a conflict.
    -> (edited without repeats, ["IssuerName on grid rows 2, 5", ...])"""
    seen: dict[tuple, tuple[str, int]] = {}
    out, bad = [], {}
    for gi, rid, ch in edited:
        key = tuple(orig[rid][k] for k in kpos)
        keep = []
        for ci in ch:
            v = rows[gi][ci]
            first = seen.setdefault((key, ci), (v, gi))
            if first[1] == gi:
                keep.append(ci)
            elif first[0] != v:
                bad.setdefault((key, ci), {first[1] + 1}).add(gi + 1)
        if keep:
            out.append((gi, rid, keep))
    return out, [f"{names[ci]} on grid rows {', '.join(map(str, sorted(g)))}"
                 for (_, ci), g in bad.items()]


def nz(s: str):
    """Grid text -> SQL value. Blank or the word NULL (any case) is SQL NULL."""
    return None if s == "" or s.strip().upper() == "NULL" else s


WRITE_WORDS = re.compile(
    r"\b(insert|update|delete|merge|into|drop|alter|create|truncate|exec|execute|grant|revoke|"
    r"deny|backup|restore|dbcc|shutdown|kill|bulk|openrowset|opendatasource|sp_\w+|xp_\w+)\b", re.I)


INT_RANGE = {"tinyint": (0, 255), "smallint": (-2 ** 15, 2 ** 15 - 1), "int": (-2 ** 31, 2 ** 31 - 1),
             "bigint": (-2 ** 63, 2 ** 63 - 1)}
TEXT_TYPES = {"char", "varchar", "nchar", "nvarchar"}
DATETIME_TYPES = {"datetime", "datetime2", "smalldatetime", "datetimeoffset"}


def is_formula(v: str) -> bool:
    """'=GETDATE()': a SQL expression the server evaluates into the cell."""
    return len(v) > 1 and v[0] == "="


def sql_type(col: "Col") -> str:
    """The column's type as written in a CAST, or "" when values are not cast to it."""
    t = col.type  # text stays uncast: the grid's own format beats CAST's 'Nov  9 2026 12:00PM'
    if t in ("decimal", "numeric"):
        return f"{t}({col.precision},{col.scale})"
    if t in ("datetime2", "datetimeoffset", "time"):
        return f"{t}({col.scale})"
    if t in INT_RANGE or t in DATETIME_TYPES or t in ("date", "bit", "float", "real", "money",
                                                      "smallmoney", "uniqueidentifier"):
        return t
    return ""


class FxError(Exception):
    """A formula that cannot be worked out (bad reference, circular, server error)."""


# Formulas: "=" + a SQL Server expression. Other cells are referenced Excel style
# (B2, $B$2, SUM(B2:B9)) or by column name (ISIN, [Issue Date]) for the same row.
FX_AGG = {"sum": "SUM", "min": "MIN", "max": "MAX", "avg": "AVG", "average": "AVG", "count": "COUNT"}
_A1 = r"\$?[A-Za-z]{1,3}\$?\d+"
_FX_TOKEN = re.compile(rf"""
    (?P<str>N?'(?:[^']|'')*')
  | (?P<agg>(?<![\w.])(?:SUM|MIN|MAX|AVG|AVERAGE|COUNT)\s*\(\s*{_A1}:{_A1}\s*\))
  | (?P<br>\[(?:[^\]]|\]\])+\])
  | (?P<cell>(?<![\w.@#$]){_A1}(?::{_A1})?(?![\w(.]))
  | (?P<word>(?<![\w@#$])[A-Za-z_@#][\w@#$]*)""", re.X | re.I)
_FX_A1 = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)")
SQL_WORDS = {"as", "and", "or", "not", "null", "is", "in", "like", "between", "case", "when", "then", "else",
             "end", "select", "from", "where", "date", "time", "day", "month", "year", "key", "type", "value"}


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def fx_names(cols: list["Col"]) -> dict[str, int]:
    """Column name / header (lower case) -> grid index; label columns excluded."""
    names: dict[str, int] = {}
    for i, c in enumerate(cols):
        if c.of is None:
            names.setdefault(c.name.lower(), i)
            names.setdefault(c.label.lower(), i)
    return names


def fx_ref_name(col: "Col") -> str:
    """How a column is written in a formula: ISIN, or [Issue Date] when it needs brackets."""
    plain = re.fullmatch(r"[A-Za-z_][\w]*", col.label) and col.label.lower() not in SQL_WORDS \
        and not re.fullmatch(_A1, col.label)
    return col.label if plain else "[" + col.label.replace("]", "]]") + "]"


def _a1(tok: str) -> tuple[int, int, bool]:
    """'$B$7' -> (column 1, row 6, row is absolute)"""
    cabs, letters, rabs, num = _FX_A1.fullmatch(tok).groups()
    return col_index(letters), int(num) - 1, bool(rabs)


def fx_parse(text: str, gi: int, cols: list["Col"], nrows: int):
    """Formula of data row `gi` -> (pieces, refs, shape). pieces: SQL text and ref
    numbers; refs: ("cell", column, row) or ("agg", FUNC, [(row, column), ...]);
    shape: the formula with references made relative to the row, equal for every
    cell that can share one query (a filled-down formula)."""
    body, names = text[1:], fx_names(cols)
    pieces: list = []
    refs: list = []
    shape: list[str] = []
    pos, prev = 0, ""

    def cell(ci: int, ri: int, row_abs: bool, tok: str) -> str:
        if not 0 <= ci < len(cols) or cols[ci].of is not None:
            raise FxError(f"{tok}: no such column")
        if not 0 <= ri < nrows:
            raise FxError(f"{tok}: row {ri + 1} does not exist")
        return f"C{ci}{'$' if row_abs else '@'}{ri if row_abs else ri - gi}"

    for m in _FX_TOKEN.finditer(body):
        gap = body[pos:m.start()]
        pieces.append(gap)
        shape.append(gap)
        pos = m.end()
        tok, kind = m.group(), m.lastgroup
        before, after = body[:m.start()].rstrip(), body[m.end():].lstrip()
        ci = None
        if kind == "br":
            ci = names.get(tok[1:-1].replace("]]", "]").lower())
        elif kind in ("word", "cell") and tok.lower() in names and not after.startswith(("(", ".")) \
                and not before.endswith(".") and prev != "as":
            ci = names[tok.lower()]
        if ci is not None:  # a column of the same row
            refs.append(("cell", ci, gi))
            pieces.append(len(refs) - 1)
            shape.append(f"\x00C{ci}@0\x00")
        elif kind == "cell":
            if ":" in tok:
                raise FxError(f"{tok}: a range works inside SUM, MIN, MAX, AVG or COUNT")
            c, r, ra = _a1(tok)
            shape.append(f"\x00{cell(c, r, ra, tok)}\x00")
            refs.append(("cell", c, r))
            pieces.append(len(refs) - 1)
        elif kind == "agg":
            fn = FX_AGG[re.match(r"\w+", tok).group().lower()]
            a, b = re.search(rf"({_A1}):({_A1})", tok).groups()
            (c0, r0, a0), (c1, r1, a1) = _a1(a), _a1(b)
            spec = f"{cell(c0, r0, a0, a)}:{cell(c1, r1, a1, b)}"
            r0, r1, c0, c1 = min(r0, r1), max(r0, r1), min(c0, c1), max(c0, c1)
            refs.append(("agg", fn, [(r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)]))
            pieces.append(len(refs) - 1)
            shape.append(f"\x00{fn}({spec})\x00")
        else:
            pieces.append(tok)
            shape.append(tok)
        prev = tok.lower() if kind == "word" else ""
    pieces.append(body[pos:])
    shape.append(body[pos:])
    return pieces, refs, "".join(shape)


def fx_shift(text: str, dr: int, dc: int, cols: list["Col"]) -> str:
    """A formula copied `dr` rows down / `dc` columns right: B2 -> B3, $B$2 stays.
    Column names refer to the same row, so they stay as they are."""
    if not (dr or dc) or not is_formula(text):
        return text
    names, body = fx_names(cols), text[1:]

    def move(m):
        cabs, letters, rabs, num = m.groups()
        ci = col_index(letters) + (0 if cabs else dc)
        ri = int(num) - 1 + (0 if rabs else dr)
        return "#REF!" if ci < 0 or ri < 0 else f"{cabs}{num2alpha(ci)}{rabs}{ri + 1}"

    out, pos = [], 0
    for m in _FX_TOKEN.finditer(body):
        tok, kind = m.group(), m.lastgroup
        out.append(body[pos:m.start()])
        pos = m.end()
        if (kind == "cell" and tok.lower() not in names) or kind == "agg":
            tok = re.sub(rf"(?<![\w$]){_FX_A1.pattern}(?!\w)", move, tok) if kind == "agg" \
                else ":".join(_FX_A1.sub(move, part) for part in tok.split(":"))
        out.append(tok)
    out.append(body[pos:])
    return "=" + "".join(out)


def fx_sql(pieces: list, refs: list, cols: list["Col"], cast: str) -> str:
    """One query for all cells of a shape: each cell is a row of the JSON parameter
    ({"i": n, "p0": value, ...}); referenced values are converted to their column's type."""
    withs, expr = ["[i] int '$.i'"], []
    for p in pieces:
        if isinstance(p, str):
            expr.append(p)
            continue
        ref = refs[p]
        if ref[0] == "cell":
            withs.append(f"[p{p}] nvarchar(max) '$.p{p}'")
            t = sql_type(cols[ref[1]])
            expr.append(f"TRY_CONVERT({t}, j.[p{p}])" if t else f"j.[p{p}]")
        else:
            withs.append(f"[p{p}] nvarchar(max) '$.p{p}' AS JSON")
            t = sql_type(cols[ref[2][0][1]]) if ref[2] else ""
            if not t and ref[1] in ("SUM", "AVG"):
                t = "float"
            v = f"TRY_CONVERT({t}, value)" if t else "value"
            expr.append(f"(SELECT {ref[1]}({v}) FROM OPENJSON(j.[p{p}]))")
    e = "".join(expr)
    v = f"CAST(({e}) AS {cast})" if cast else f"({e})"
    return f"SELECT j.[i], {v} AS v FROM OPENJSON(?) WITH ({', '.join(withs)}) j"


def cell_error(col: "Col", v: str) -> str | None:
    """Why `v` cannot go into `col`, or None when it can. Blank / NULL always can."""
    if col.external or col.readonly or col.audit or is_formula(v):  # formulas: checked once evaluated
        return None
    if col.big and v and v.strip().upper() != "NULL":  # unknown values: see App._big_lookups
        if v in col.nomatch:
            return col.nomatch[v]
        if v in col.missing:
            return f"Not in the lookup ({src_name(col)})"
        if v not in col.key_fields:
            return None
    if col.choices is not None and v not in col.choices and v != "" and v.strip().upper() != "NULL":
        return f"Not in the lookup list ({col.lookup})"
    return _value_error(col.type, col.size, col.precision, col.scale, v)


def col_info(col: "Col") -> str:
    """'ISIN · varchar(12) · primary key · lookup: foreign key to dbo.Bond.ISIN'"""
    if col.scratch:
        return f"{col.label} · scratch column: on this sheet only, never written to the database"
    if col.external:
        return f"{col.label} · not from the target table, read-only"
    t = col.type
    if t in TEXT_TYPES:
        t += f"({'max' if col.size == -1 else col.size})"
    elif t in ("decimal", "numeric"):
        t += f"({col.precision},{col.scale})"
    parts = [col.name, t]
    if col.is_pk:
        parts.append("primary key")
    if col.readonly:
        parts.append("read-only")
    elif col.audit:
        parts.append("set automatically")
    if col.big:
        parts.append(f"lookup {col.lookup}: labels read per value")
    elif col.choices is not None:
        parts.append(f"{len(col.choices):,} allowed values from {col.lookup}")
    elif col.fk:
        parts.append(f"foreign key to {col.fk}")
    return "  ·  ".join(parts)


def load_lookup_file(path: str) -> dict:
    """Lookup file: JSON object, key = column, value = the allowed values or a query.

        {"Sales.dbo.Orders.Status": ["Open", "Shipped", "Closed"],
         "dbo.Orders.Currency": {"sql": "SELECT Code FROM dbo.Currency", "database": "Ref"},
         "CountryCode": {"sql": "SELECT Code FROM dbo.Country"}}

    A key names a column as database.schema.table.column, schema.table.column,
    table.column or just column (any table); the most specific key wins."""
    if not path:
        return {}
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Lookup file must be a JSON object")
    return {k.lower().replace("[", "").replace("]", ""): v for k, v in data.items()}


def lookup_for(spec: dict, database: str, table: str, column: str):
    schema, name = split_table(table)
    for key in (f"{database}.{schema}.{name}.{column}", f"{schema}.{name}.{column}", f"{name}.{column}",
                column):
        if key.lower() in spec:
            return spec[key.lower()]
    return None


NO_LOOKUP = "-"  # saved label choice: no dropdown for this foreign key


def guess_label(names: list[str], table: str = "") -> str:
    """The text column that most likely names a lookup row: *Name, then one named
    like its table (Sector in lookup.Sector), then *Description, then *Code, then
    the first; audit columns never."""
    names = [n for n in names if not re.search(r"(created|modified|updated)(by|user)$", n, re.I)]
    own = re.escape(split_table(table)[1].rstrip("s")) if table else "(?!)"
    for pat in (r"name$", rf"^{own}s?$", r"desc(ription)?$", r"code$"):
        hit = [n for n in names if re.search(pat, n, re.I)]
        if hit:
            return hit[0]
    return names[0] if names else ""


def choice_search(search_for: str, data) -> int | None:
    """Dropdown search: which entry ("key" or "key  ·  label") the typed text
    means. Exact key, exact label, key or label starting with it, label
    containing it - in that order, ignoring case. No fuzzy guesses."""
    t = search_for.strip().lower()
    if not t:
        return None
    split = [(i, *(str(e).lower().split("  ·  ", 1) + [""])[:2]) for i, e in enumerate(data) if str(e)]
    for test in (lambda k, lb: k == t, lambda k, lb: lb == t, lambda k, lb: k.startswith(t),
                 lambda k, lb: lb.startswith(t), lambda k, lb: t in lb):
        for i, k, lb in split:
            if test(k, lb):
                return i
    return None


def resolve_choice(col: "Col", v: str) -> str:
    """Typed text in a lookup column -> the key it means (unchanged if none)."""
    if col.choices is None or v in col.choices or v == "" or v.strip().upper() == "NULL":
        return v
    i = choice_search(v, col.choice_list)
    if i is None:
        return v
    entry = col.choice_list[i]
    return col.picked.get(entry, entry)


LABEL_MODES = {"Same column (6 · Valoren)": "same", "Value only (Valoren)": "value",
               "Separate column": "separate", "Off (id only)": "off"}
LABEL_FORMATS = {"6 · Valoren": "{id} · {label}", "Valoren (6)": "{label} ({id})", "Valoren": "{label}"}
MISMATCH = {"⚠ before the value": "mark", "Red text": "red", "Both": "both"}
WARN = "⚠ "


LABEL_SEP = " · "  # between label fields; Settings can change it (set_label_sep)
LABEL_SEPS = [" · ", " - ", " | ", ", ", " / "]


def set_label_sep(sep: str) -> None:
    global LABEL_SEP
    LABEL_SEP = sep or " · "


def label_format_error(fmt: str) -> str | None:
    """Why a label format cannot be used, or None. It needs {label}; {id} is optional."""
    if "{label}" not in fmt:
        return "The format needs {label}, e.g. {id} - {label}"
    try:
        fmt.format(id="6", label="Valoren")
    except (KeyError, IndexError, ValueError):
        return "Only {id} and {label} can be used; write {{ or }} for a brace"
    return None


def label_fields(src: dict) -> list[str]:
    """The label fields of a table lookup; saved as one name (older settings) or a list."""
    label = src.get("label")
    return [label] if isinstance(label, str) and label else list(label or [])


def src_name(col: "Col") -> str:
    src = col.src or {}
    return f"{src.get('db')}.{src.get('table')}.{src.get('key')}"


def field_text(col: "Col", part: int, fv: str, mark: bool) -> str:
    """Shown text of label field `part` holding `fv`: through its own lookup if it has one."""
    sub = col.subs.get(part)
    if sub is None or not fv:
        return fv
    known = sub.in_lookup(fv)
    if known:
        label = label_of(sub, fv, mark)
        return sub.fmt.format(id=fv, label=label) if label and sub.fmt else fv
    return WARN + fv if known is False and mark else fv


def label_of(col: "Col", v: str, mark: bool) -> str:
    """All label fields of key `v`, joined; fields with their own lookup show its label."""
    vals = col.key_fields.get(v)
    if vals is None:
        return col.key_label.get(v, "")
    if not col.subs:  # joined now, so a new separator shows at once
        return LABEL_SEP.join(str(x) for x in vals if x not in (None, ""))
    return LABEL_SEP.join(t for p, fv in enumerate(vals) if (t := field_text(col, p, str(fv), mark)))


def label_text(col: "Col", v: str, mark: bool, part: int | None = None) -> str:
    """Text of a label column for value `v` of lookup column `col`: all fields,
    or field number `part`."""
    if not v:
        return ""
    known = col.in_lookup(v)
    if known is None:  # big lookup, not read yet
        return ""
    if known:
        if part is None:
            return label_of(col, v, mark)
        vals = col.key_fields.get(v, ())
        return field_text(col, part, str(vals[part]), mark) if part < len(vals) else ""
    return f"{WARN}not in lookup" if mark else "not in lookup"


def lookup_sql(table: str, key: str, labels: list[str] = ()) -> str:
    """Values of a lookup table's key column, with the `labels` fields beside them."""
    schema, name = split_table(table)
    if not labels:
        return (f"SELECT TOP ({MAX_CHOICES + 1}) {q(key)} FROM {q(schema)}.{q(name)} "
                f"WHERE {q(key)} IS NOT NULL GROUP BY {q(key)} ORDER BY 1")
    cols = ", ".join(q(c) for c in (key, *labels))
    # one row per key; ROW_NUMBER, not MIN(), so a field of any type (bit, ...) works
    return (f"SELECT TOP ({MAX_CHOICES + 1}) {cols} FROM (SELECT {cols}, ROW_NUMBER() OVER "
            f"(PARTITION BY {q(key)} ORDER BY (SELECT 0)) AS [rn·] FROM {q(schema)}.{q(name)} "
            f"WHERE {q(key)} IS NOT NULL) x WHERE [rn·] = 1 ORDER BY 1")


def col_key(database: str, table: str, column: str) -> str:
    """Settings key of a column: database.schema.table.column, lower case."""
    return ".".join((database, *split_table(table), column)).lower()


def set_choices(col: "Col", rows: list, source: str, fields: list[str] = ()) -> None:
    """rows: values, or [value, label, ...] (the other columns are shown beside the
    value). `fields` names those columns."""
    rows = [r if isinstance(r, (list, tuple)) else [r] for r in rows]
    rows = list({str(r[0]): r for r in rows if r and str(r[0]) != ""}.values())
    if len(rows) > MAX_CHOICES:  # too many to list: leave it to the server
        return
    width = max((len(r) for r in rows), default=1) - 1
    col.label_fields = list(fields)[:width] or (["label"] if width else [])
    col.label_fields += [f"label {i + 1}" for i in range(len(col.label_fields), width)]
    col.choices, col.lookup = frozenset(str(r[0]) for r in rows), source
    col.choice_list, col.picked, col.key_label, col.key_fields = [], {}, {}, {}
    for r in rows:
        key = str(r[0])
        vals = tuple("" if v is None else str(v) for v in r[1:]) + ("",) * (width - len(r) + 1)
        label = LABEL_SEP.join(v for v in vals if v)
        entry = f"{key}  ·  {label}" if label else key
        col.choice_list.append(entry)
        if label:
            col.picked[entry], col.key_label[key], col.key_fields[key] = key, label, vals


@functools.lru_cache(maxsize=65536)
def _value_error(t: str, size: int, precision: int, scale: int, v: str) -> str | None:
    if v == "" or v.strip().upper() == "NULL":
        return None
    if t in INT_RANGE:
        if not re.fullmatch(r"-?\d+", v):
            return "Whole number expected"
        lo, hi = INT_RANGE[t]
        return None if lo <= int(v) <= hi else f"Out of range for {t}"
    if t in ("decimal", "numeric"):
        m = re.fullmatch(r"-?(\d*)(?:\.(\d*))?", v)
        if not m or not (m.group(1) or m.group(2)):
            return "Number expected"
        whole, frac = m.group(1).lstrip("0"), (m.group(2) or "").rstrip("0")
        if len(whole) > precision - scale or len(frac) > scale:
            return f"Too many digits for {t}({precision},{scale})"
        return None
    if t in ("float", "real", "money", "smallmoney"):
        try:
            float(v)
            return None
        except ValueError:
            return "Number expected"
    if t == "bit":
        return None if v.lower() in ("0", "1", "true", "false") else "Use 0 or 1"
    if t == "date":
        try:
            dt.date.fromisoformat(v)
            return None
        except ValueError:
            return "Date as YYYY-MM-DD expected"
    if t in DATETIME_TYPES:
        try:
            dt.datetime.fromisoformat(v)
            return None
        except ValueError:
            return "Date and time as YYYY-MM-DD hh:mm:ss expected"
    if t == "time":
        try:
            dt.time.fromisoformat(v)
            return None
        except ValueError:
            return "Time as hh:mm:ss expected"
    if t == "uniqueidentifier":
        return None if re.fullmatch(r"\{?[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}\}?", v) \
            else "GUID expected"
    if t in TEXT_TYPES and size > 0 and len(v) > size:
        return f"Longer than {size} characters"
    return None


NUMERIC = {"tinyint", "smallint", "int", "bigint", "decimal", "numeric", "float", "real", "money",
           "smallmoney", "bit"}


def in_clause(text: str, cols: dict[str, "Col"]) -> str | None:
    """'Col\nv1\nv2' (first spreadsheet column) -> "[Col] IN ('v1', 'v2')".
    None when the text is not a header plus values. Numbers stay unquoted for
    numeric columns; repeats are dropped."""
    lines = [ln.split("\t")[0].strip() for ln in text.replace("\r", "").split("\n")]
    lines = [ln for ln in lines if ln]
    if len(lines) < 2:
        return None
    head = lines[0].strip("[]")
    col = cols.get(head.lower())
    if col is None and not re.fullmatch(r"[A-Za-z_]\w{0,127}", head):  # unknown names: plain words only
        return None
    name = col.name if col else head
    values = list(dict.fromkeys(lines[1:]))
    numeric = col is not None and col.type in NUMERIC and all(
        re.fullmatch(r"-?\d+(\.\d+)?", v) for v in values)
    vals = values if numeric else ["'" + v.replace("'", "''") + "'" for v in values]
    return f"{q(name)} IN ({', '.join(vals)})"


def edit_target(db: "Db", database: str, table: str, srcs) -> tuple[str, str, list["Col"]]:
    """Which table a query result edits: (database, schema.table, its columns), or
    ("", "", []) for a read-only result. The table picked in the toolbar when the
    result has its columns, else the first table (in column order) whose key
    columns are all in the result, else the first table that resolves."""
    found = list(dict.fromkeys(s[:3] for s in srcs or [] if s))
    if not found:
        return "", "", []
    picked = [f for f in found if f[0].lower() == database.lower()
              and (f[1].lower(), f[2].lower()) == tuple(x.lower() for x in split_table(table))] if table else []
    first = None
    for tdb, schema, name in picked or found:
        try:
            cols = db.columns(tdb, f"{schema}.{name}")
        except Exception:  # noqa: BLE001 - e.g. a system view: not editable
            continue
        if first is None:
            first = (tdb, f"{schema}.{name}", cols)
        shown = {s[3].lower() for s in srcs if s and s[:3] == (tdb, schema, name)}
        keys = [c.name.lower() for c in cols if c.is_pk]
        if picked or (keys and all(k in shown for k in keys)):
            return tdb, f"{schema}.{name}", cols
    return first or ("", "", [])


def table_refs(sql: str, table: str) -> int:
    """How many times `table` is named after FROM / JOIN / a comma in a FROM list."""
    bare = re.sub(r"--[^\n]*|/\*.*?\*/", " ", sql, flags=re.S)
    bare = re.sub(r"'(?:[^']|'')*'", "''", bare)
    want = tuple(x.lower() for x in split_table(table))
    name = r"((?:\[[^\]]+\]|[\w#]+)(?:\s*\.\s*(?:\[[^\]]+\]|[\w#]+)?){0,3})"
    hits = 0
    for m in re.finditer(r"(?:\bfrom|\bjoin|,)\s+" + name, bare, re.I):
        parts = [p.strip().strip("[]").lower() for p in m.group(1).split(".")]
        got = (parts[-2] if len(parts) > 1 and parts[-2] else "dbo", parts[-1])
        hits += got == want
    return hits


def check_select(sql: str) -> str | None:
    """Return an error message unless `sql` is a single read-only SELECT / WITH."""
    bare = re.sub(r"--[^\n]*|/\*.*?\*/", " ", sql, flags=re.S)
    bare = re.sub(r"'(?:[^']|'')*'", "''", bare)        # string literals
    bare = re.sub(r"\[(?:[^\]]|\]\])*\]", "[x]", bare)  # bracketed names
    bare = bare.strip().rstrip(";").strip()
    if not bare:
        return "Query is empty"
    if not re.match(r"(select|with)\b", bare, re.I):
        return "Query must start with SELECT or WITH"
    if ";" in bare:
        return "One statement only"
    m = WRITE_WORDS.search(bare)
    if m:
        return f"Read-only queries only: '{m.group(0)}' is not allowed"
    return None


# ── UI ───────────────────────────────────────────────────────────────────────
class Tab:
    """One sheet tab: its own connection, source, grid and status message.
    App reads and writes these through properties (see TAB_STATE), so App code
    always works on the tab in `App.tab`."""

    def __init__(self, n: int):
        self.n = n
        self.db = Db()
        self.cols: list[Col] = []
        self.keys: list[Col] = []
        self.database = self.table = ""
        self.orig: dict[str, list[str]] = {}  # rid -> original values (loaded rows)
        self.table_cols: list[Col] = []  # every column of the target table
        self.from_query = False
        self._all_tables: list[str] = []
        self.filters: dict[str, dict] = {}
        self._drag: dict | None = None
        self._near_handle = False
        self._busy = False
        self.sheet: Sheet | None = None
        self.conn_text = "Not connected"
        self.conn_name: str | None = None  # saved connection this tab uses
        self.server = self.uid = self.pwd = ""
        self.tried = False  # auto-connect already attempted
        self.dbs: list[str] = []
        self.ui = {"db": "", "table": "", "where": "", "top": "1000", "sql": "", "dry": False}
        self.sql_open = False
        self.note = ("Ready", False)
        self.job: tuple | None = None  # (label, started, cancellable) while a job runs
        self.bad: dict[tuple[int, int], str] = {}  # (data row, column) -> why the value is invalid
        self.fx: dict[int, tuple[list, dict[int, dict]]] = {}  # id(row) -> (row, {column: formula})
        self.fx_why: dict[tuple[int, int], str] = {}  # formula cells not calculated, and why
        self.fx_hold = False  # calculation cancelled or failed: wait for the next entered formula
        self.lk_hold = False  # same for reading big lookups
        self.label: tk.Label | None = None  # its tab in the tab strip

    @property
    def title(self) -> str:
        return self.table or ("Query" if self.cols else self.ui["table"] or f"Sheet{self.n}")


TAB_STATE = ("db", "cols", "keys", "database", "table", "orig", "table_cols", "from_query",
             "_all_tables", "filters", "_drag", "_near_handle", "_busy", "sheet", "conn_text", "bad")


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.user = getpass.getuser()
        self.jobs: queue.Queue = queue.Queue()
        self._mark_after = None
        self._fx_cell = None  # (tab, data row, Col, text shown) the fx bar is editing
        self.log: list[tuple[str, str, str, bool]] = []  # (time, sheet, text, error), see msg
        self._abandoned: dict[object, Db] = {}  # job token -> connection replaced after Cancel
        self._running: set[object] = set()  # tokens of jobs whose worker has not returned
        root.report_callback_exception = self._uncaught
        self.tabs: list[Tab] = []
        self.tab: Tab | None = None  # tab App code works on; differs from shown only
        self.shown: Tab | None = None  # while a background tab's job finishes (_in)
        self._tab_seq = 0

        root.title(APP_NAME)
        root.geometry("1440x860")
        root.minsize(1100, 520)
        self.settings = load_settings()
        set_label_sep(self.settings.get("label_sep", " · "))
        if not self.settings["ime"]:
            # X11 input methods (ibus, fcitx) cost an IM-server round trip per ttk
            # widget, ~60 ms each under XWayland: seconds at start and on every tab
            # switch. Off unless the IME setting is on (needed only for composed input).
            root.tk.call("tk", "useinputmethods", 0)
        self._painted: list = []
        C.update(DARK if self.settings["theme"] == "Dark" else LIGHT)
        root.configure(bg=C["bg"])
        self._style()
        self._build()
        self._restore_tabs()
        root.protocol("WM_DELETE_WINDOW", self.quit)
        self._poll_id = root.after(50, self._poll)

    # look
    def _style(self):
        s = ttk.Style(self.root)
        s.theme_use("clam")
        s.configure(".", font=FONT, background=C["surface"], foreground=C["text"],
                    bordercolor=C["line"], lightcolor=C["surface"], darkcolor=C["surface"],
                    focuscolor=C["accent"], troughcolor=C["bg"], selectbackground=C["accent_soft"],
                    selectforeground=C["text"], insertcolor=C["text"])
        s.configure("TFrame", background=C["surface"])
        s.configure("Bar.TFrame", background=C["bg"])
        s.configure("TLabel", background=C["surface"], foreground=C["text2"])
        s.configure("Group.TLabel", background=C["surface"], foreground=C["muted"])
        s.configure("Status.TLabel", background=C["bg"], foreground=C["text2"])
        s.configure("Err.TLabel", background=C["bg"], foreground=C["error"])
        self.root.option_add("*TCombobox*Listbox.background", C["surface"])
        self.root.option_add("*TCombobox*Listbox.foreground", C["text"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", C["accent_soft"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", C["text"])
        s.configure("TCheckbutton", background=C["surface"], foreground=C["text"],
                    indicatorbackground=C["surface"], indicatorforeground=C["text"])
        s.map("TCheckbutton", background=[("active", C["surface"])],
              indicatorbackground=[("selected", C["button"])],
              indicatorforeground=[("selected", "#ffffff")])
        s.configure("Treeview", background=C["surface"], fieldbackground=C["surface"], foreground=C["text"],
                    bordercolor=C["line"], rowheight=24, font=FONT)
        s.map("Treeview", background=[("selected", C["accent_soft"])], foreground=[("selected", C["text"])])
        s.configure("Treeview.Heading", background=C["header"], foreground=C["text2"], font=FONT,
                    bordercolor=C["line"], relief="flat")
        s.map("Treeview.Heading", background=[("active", C["hover"])])
        s.configure("Vertical.TScrollbar", background=C["bg"], troughcolor=C["surface"],
                    bordercolor=C["surface"], arrowcolor=C["text2"])
        s.configure("Horizontal.TScrollbar", background=C["bg"], troughcolor=C["surface"],
                    bordercolor=C["surface"], arrowcolor=C["text2"])

        # Rounded buttons and fields: ttk cannot round corners, so each state is a small
        # anti-aliased image stretched around the content (9-slice). Element names are
        # per theme because ttk elements cannot be redefined.
        self._gen = getattr(self, "_gen", 0) + 1
        self._imgs: list[tk.PhotoImage] = []  # ttk only keeps image names

        def box(fill, border):
            img = rounded(24, 24, 6, C[fill], C[border], C["surface"])
            self._imgs.append(img)
            return img

        def element(name, normal, *states):
            full = f"aq{self._gen}.{name}"
            s.element_create(full, "image", normal, *states, border=7, padding=0, sticky="nsew")
            return full

        ghost = element("ghost", box("surface", "surface"), ("disabled", box("surface", "surface")),
                        ("pressed", box("header_sel", "header_sel")),
                        ("selected", box("accent_soft", "accent")), ("active", box("hover", "hover")))
        outline = element("outline", box("surface", "field"), ("disabled", box("surface", "line")),
                          ("pressed", box("header_sel", "field")), ("active", box("hover", "field")))
        primary = element("primary", box("button", "button"), ("disabled", box("bg", "line")),
                          ("pressed", box("button_hover", "button_hover")),
                          ("active", box("button_hover", "button_hover")))
        field = element("field", box("surface", "field"), ("disabled", box("bg", "line")),
                        ("focus", box("surface", "accent")))
        arrow = f"aq{self._gen}.arrow"
        s.element_create(arrow, "image", self.icon("chevron-down"),
                         ("disabled", self.icon("chevron-down", "disabled")), sticky="", padding=(0, 0, 6, 0))

        def button_layout(bg, part="Button"):
            return [(bg, {"sticky": "nsew", "children": [
                (f"{part}.padding", {"sticky": "nsew", "children": [
                    (f"{part}.label", {"sticky": "nsew"})]})]})]

        for style, bg, fg in (("TButton", outline, C["text"]), ("Tool.TButton", ghost, C["text"]),
                              ("Primary.TButton", primary, "#ffffff")):
            s.layout(style, button_layout(bg))
            s.configure(style, padding=(10, 6), foreground=fg, background=C["surface"], anchor="center",
                        width=0)
            s.map(style, foreground=[("disabled", C["disabled"])])
        for style, part in (("Toggle.TCheckbutton", "Toolbutton"), ("Tool.TMenubutton", "Menubutton")):
            s.layout(style, button_layout(ghost, part))
            s.configure(style, padding=(10, 6), foreground=C["text"], background=C["surface"], width=0)
            s.map(style, foreground=[("disabled", C["disabled"])])

        s.layout("TEntry", [(field, {"sticky": "nsew", "children": [
            ("Entry.padding", {"sticky": "nsew", "children": [("Entry.textarea", {"sticky": "nsew"})]})]})])
        s.layout("TCombobox", [(field, {"sticky": "nsew", "children": [
            (arrow, {"side": "right", "sticky": ""}),
            ("Combobox.padding", {"sticky": "nsew", "children": [
                ("Combobox.textarea", {"sticky": "nsew"})]})]})])
        for w in ("TEntry", "TCombobox"):
            s.configure(w, padding=(8, 5), foreground=C["text"], fieldbackground=C["surface"],
                        background=C["surface"])
            s.map(w, foreground=[("disabled", C["disabled"]), ("readonly", C["text"])],
                  fieldbackground=[("readonly", C["surface"])])

    def icon(self, name: str, role: str = "text2") -> tk.PhotoImage:
        """16px line icon in a palette color (text2, disabled) or "white"."""
        color = "#ffffff" if role == "white" else C[role]
        cache = self.__dict__.setdefault("_icons", {})
        if (name, color) not in cache:
            cache[name, color] = tk.PhotoImage(data=ICONS[name][color], format="png")
        return cache[name, color]

    def with_icon(self, widget, name: str, white=False):
        """Give a ttk button an icon (and a disabled version); re-applied on theme change."""
        self.__dict__.setdefault("_iconed", []).append((widget, name, white))
        widget.configure(image=(self.icon(name, "white" if white else "text2"),
                                "disabled", self.icon(name, "disabled")), compound="left")
        return widget

    def paint(self, w, **opts):
        """Create-time colors for plain tk widgets; re-applied on theme change.
        opts map a widget option to a palette key, e.g. bg="surface"."""
        self._painted.append((w, opts))
        w.configure(**{k: C[v] for k, v in opts.items()})
        return w

    def apply_theme(self):
        C.clear()
        C.update(DARK if self.settings["theme"] == "Dark" else LIGHT)
        self.v_set["theme"].set(self.settings["theme"])
        self.root.configure(bg=C["bg"])
        self._style()
        for w, name, white in self._iconed:
            w.configure(image=(self.icon(name, "white" if white else "text2"),
                               "disabled", self.icon(name, "disabled")))
        for ib in (self.i_conn, self.i_settings):
            ib.redraw()
        for w, opts in self._painted:
            try:
                w.configure(**{k: C[v] for k, v in opts.items()})
            except tk.TclError:  # widget destroyed
                pass
        for t in self.tabs:
            t.sheet.set_options(**sheet_colors())
        self._draw_tabs()
        self._marks()
        self._place_handle()

    def _toggle_setting(self, key: str):
        self.settings[key] = self.v_set[key].get()
        save_settings(self.settings)
        if key == "theme":
            self.apply_theme()

    def _build(self):
        load_env()
        self.v_server, self.v_uid, self.v_pwd = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.v_db, self.v_table, self.v_where = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.v_top = tk.StringVar(value="1000")
        self.v_dry = tk.BooleanVar(value=False)
        self.v_sql_open = tk.BooleanVar(value=False)
        self.v_set = {"review": tk.BooleanVar(value=self.settings["review"]),
                      "theme": tk.StringVar(value=self.settings["theme"])}

        # one toolbar: source, filter, rows, push, then connection and settings
        bar = ttk.Frame(self.root, padding=(12, 10))
        bar.pack(fill="x")
        self.paint(tk.Frame(self.root, height=1), bg="line").pack(fill="x")

        def divider(side="left"):
            self.paint(tk.Frame(bar, width=1, height=20), bg="line").pack(side=side, padx=12)

        # right side first, so the Where field takes whatever width is left
        self.i_settings = IconButton(self, bar, "settings", lambda: SettingsDialog(self), "Settings")
        self.i_settings.pack(side="right")
        self.i_conn = IconButton(self, bar, "server", lambda: ConnDialog(self),
                                 lambda: self.conn_text, dot=True)
        self.i_conn.pack(side="right", padx=(0, 4))
        divider("right")
        self.b_update = self.with_icon(ttk.Button(bar, text="Update", style="Primary.TButton",
                                                  command=self.update), "pencil", white=True)
        self.b_update.pack(side="right")
        self.b_insert = self.with_icon(ttk.Button(bar, text="Insert", style="Primary.TButton",
                                                  command=self.insert), "plus", white=True)
        self.b_insert.pack(side="right", padx=(0, 8))
        self.with_icon(ttk.Checkbutton(bar, text="Dry run", style="Toggle.TCheckbutton",
                                       variable=self.v_dry), "flask-conical").pack(side="right", padx=(0, 8))

        self.cb_db = ttk.Combobox(bar, textvariable=self.v_db, width=14, state="readonly")
        self.cb_db.pack(side="left")
        # a database picked by hand starts with no table: the old one belongs elsewhere
        self.cb_db.bind("<<ComboboxSelected>>", lambda e: (self.v_table.set(""), self.pick_db()))
        Hint(self, self.cb_db, self.v_db, "Database")
        self.cb_table = ttk.Combobox(bar, textvariable=self.v_table, width=22)
        self.cb_table.pack(side="left", padx=(8, 0))
        self.cb_table.bind("<KeyRelease>", self._filter_tables)
        self.cb_table.bind("<Return>", lambda e: self.load())
        Hint(self, self.cb_table, self.v_table, "Table")
        divider()
        e_where = ttk.Entry(bar, textvariable=self.v_where, font=MONO)
        e_where.pack(side="left", fill="x", expand=True)
        e_where.bind("<Return>", lambda e: self.load())
        e_where.bind("<<Paste>>", self._paste_where)
        Hint(self, e_where, self.v_where, "Where")
        ttk.Label(bar, text="Top", style="Group.TLabel").pack(side="left", padx=(12, 6))
        digits = (self.root.register(lambda v: v.isdigit() or v == ""), "%P")
        ttk.Entry(bar, textvariable=self.v_top, width=7, justify="right", validate="key",
                  validatecommand=digits).pack(side="left")
        self.b_load = self.with_icon(ttk.Button(bar, text="Load", style="Tool.TButton",
                                                command=self.load), "play")
        self.b_load.pack(side="left", padx=(8, 0))
        self.b_sql = self.with_icon(ttk.Checkbutton(bar, text="SQL", style="Toggle.TCheckbutton",
                                                    variable=self.v_sql_open, command=self.toggle_sql),
                                    "code")
        self.b_sql.pack(side="left", padx=(4, 0))
        divider()
        rows = self.with_icon(ttk.Menubutton(bar, text="Rows", style="Tool.TMenubutton"), "rows-3")
        rows.pack(side="left")
        menu = self.paint(tk.Menu(rows, tearoff=0, relief="flat", bd=1, font=FONT),
                          bg="surface", fg="text", activebackground="hover", activeforeground="text",
                          disabledforeground="disabled")
        rows.configure(menu=menu)
        self.b_blank = MenuItem(menu, "Blank grid", self.blank)
        self.b_paste = MenuItem(menu, "Paste with headers", self.paste_with_headers)
        menu.add_separator()
        self.b_import = MenuItem(menu, "Import file…", self.import_file)
        self.b_export = MenuItem(menu, "Export CSV…", self.export_csv)

        # free SQL panel, hidden until toggled
        self.sql_box = ttk.Frame(self.root, padding=(12, 8))
        sql_wrap = self.paint(tk.Frame(self.sql_box, padx=1, pady=1), bg="line")
        sql_wrap.pack(side="left", fill="both", expand=True, padx=(0, 8))
        self.t_sql = self.paint(tk.Text(sql_wrap, height=5, font=MONO, wrap="none", undo=True,
                                        relief="flat", padx=8, pady=6, highlightthickness=0),
                                bg="surface", fg="text", insertbackground="text",
                                selectbackground="accent_soft", selectforeground="text")
        self.t_sql.pack(fill="both", expand=True)
        self.t_sql.bind("<Control-Return>", lambda e: (self.run_query(), "break")[1])
        side = ttk.Frame(self.sql_box)
        side.pack(side="left", fill="y")
        self.b_run = self.with_icon(ttk.Button(side, text="Run", command=self.run_query), "play")
        self.b_run.pack(fill="x")
        ttk.Label(side, text="Ctrl+Enter", style="Group.TLabel").pack(anchor="w", pady=(6, 0))

        # name box + formula bar
        fx = ttk.Frame(self.root, padding=(12, 6))
        fx.pack(fill="x")
        self.fx_bar = fx
        self.v_name = tk.StringVar()
        ttk.Entry(fx, textvariable=self.v_name, width=22, state="readonly").pack(side="left")
        self.paint(tk.Label(fx, text="fx", font=MONO), bg="surface", fg="muted").pack(side="left", padx=10)
        self.v_fx = tk.StringVar()
        self.e_fx = ttk.Entry(fx, textvariable=self.v_fx, font=CELL)
        self.e_fx.pack(side="left", fill="x", expand=True)
        self.e_fx.bind("<Return>", self._fx_commit)
        self.e_fx.bind("<FocusOut>", lambda e: self._fx_flush())
        self.e_fx.bind("<Escape>", lambda e: (self._fx_revert(), self.sheet.focus_set()))

        self.paint(tk.Frame(self.root, height=1), bg="line").pack(fill="x")
        self.grid_wrap = self.paint(tk.Frame(self.root), bg="line")
        self.grid_wrap.pack(fill="both", expand=True)

        # sheet tab strip (tabs are drawn by _draw_tabs)
        tabs = self.paint(tk.Frame(self.root), bg="bg")
        tabs.pack(fill="x")
        self.tab_bar = self.paint(tk.Frame(tabs), bg="bg")
        self.tab_bar.pack(side="left", padx=(48, 0))
        self.l_status = self.paint(tk.Label(tabs, text="", font=FONT), bg="bg", fg="text2")
        self.l_status.pack(side="right", padx=12)

        # status bar
        self.paint(tk.Frame(self.root, height=1), bg="line").pack(fill="x")
        foot = ttk.Frame(self.root, style="Bar.TFrame", padding=(12, 3))
        foot.pack(fill="x")
        self.b_cancel = self.with_icon(ttk.Button(foot, text="Cancel", style="Tool.TButton",
                                                  command=self.cancel_job), "x")
        self.l_stats = ttk.Label(foot, text="", style="Status.TLabel")
        self.l_stats.pack(side="right", padx=(16, 0))
        self.l_msg = ttk.Label(foot, text="Ready", style="Status.TLabel", cursor="hand2")
        self.l_msg.pack(side="left", fill="x", expand=True)
        self.l_msg.bind("<Button-1>", lambda e: MessageLog(self))

        self.busy_buttons = [self.b_load, self.b_blank, self.b_insert,
                             self.b_update, self.b_paste, self.b_import, self.b_run]
        self._set_table_state(False)
        self.root.bind("<Control-Return>", lambda e: self.load())
        tab_keys = {"<Control-t>": self.add_tab, "<Control-w>": self.close_tab,
                    "<Control-Next>": lambda: self.cycle_tab(1),
                    "<Control-Prior>": lambda: self.cycle_tab(-1)}
        for seq, fn in tab_keys.items():
            self.root.bind(seq, lambda e, fn=fn: (fn(), "break")[1])
        # Text binds Ctrl+T (transpose) and Ctrl+PgUp/PgDn itself, which would stop the root binding
        for seq, fn in tab_keys.items():
            self.root.bind_class("Text", seq, lambda e, fn=fn: (fn(), "break")[1])
        # Tk on Linux maps Ctrl+A to "start of line"; make it select all, as on Windows
        for cls in ("Text", "TEntry", "TCombobox", "Entry"):
            for seq in ("<Control-a>", "<Control-A>"):
                self.root.bind_class(cls, seq, select_all)
            self.root.bind_class(cls, "<<Paste>>", paste_replace)

    def _new_sheet(self) -> Sheet:
        """Grid for the tab in self.tab (not packed; activate() shows it)."""
        # tksheet flushes Tk's idle queue four times while building, which lays out
        # the whole window each time; the next normal redraw does it once instead
        idle, tk.Misc.update_idletasks = tk.Misc.update_idletasks, lambda w: None
        try:
            sheet = Sheet(
                self.grid_wrap, data=[], headers=[], font=(CELL[0], CELL[1], "normal"),
                header_font=(CELL[0], CELL[1], "normal"), index_font=(CELL[0], CELL[1], "normal"),
                paste_can_expand_x=False, paste_can_expand_y=True, show_x_scrollbar=True,
                default_row_height=22, default_column_width=110, default_row_index_width=48,
                header_align="center", index_align="center", find_bindings=["<Control-f>"], **sheet_colors(),
            )
        finally:
            tk.Misc.update_idletasks = idle
        sheet.enable_bindings("all")
        # sort is offered as whole-row sort from the header menu; the per-cell sorts
        # would scramble rows, and columns come from the table so cannot be added
        sheet.disable_bindings("sort_cells", "sort_row", "sort_column", "sort_rows",
                               "sort_columns", "insert_columns", "delete_columns")
        sheet.bind("<<SheetModified>>",
                   lambda e: (self._sync_columns(), self._schedule_marks(), self._on_select()))
        sheet.bind("<<SheetSelect>>", lambda e: (self._on_select(), self._place_handle()))
        sheet.edit_validation(self._validate_edit)
        sheet.extra_bindings("begin_edit_cell", self._begin_edit)
        sheet.disp = {}  # data column -> fn(stored text) -> shown text, see _label_display
        raw_str = sheet.MT.cell_str

        def cell_str(datarn, datacn, get_displayed=False, **kw):
            text = raw_str(datarn, datacn, get_displayed, **kw)
            fn = sheet.disp.get(datacn) if get_displayed else None
            return fn(text) if fn else text

        sheet.MT.cell_str = cell_str
        raw_find = sheet.MT.find_match

        def find_match(find, r, c):  # Ctrl+F: the stored id or the label shown with it
            if raw_find(find, r, c):
                return True
            fn, col = sheet.disp.get(c), self.cols[c] if c < len(self.cols) else None
            if not find or (fn is None and not (col and col.labelled)):
                return False
            try:
                v = sheet.MT.data[r][c]
            except IndexError:
                return False
            v = "" if v is None else str(v)
            shown = fn(v) if fn else label_of(col, v, False)
            return find in shown.lower()

        sheet.MT.find_match = find_match
        for w in (sheet.MT, sheet.CH, sheet.RI):
            w.bind("<Control-F>", lambda e: (self.find_column(), "break")[1])
        sheet.MT.bind("<Motion>", self._hover_cell, add="+")
        sheet.CH.bind("<Motion>", self._hover_header, add="+")
        self.sheet = sheet
        self._setup_excel()
        self._scratch()
        return sheet

    def _scratch(self):
        """Empty sheet before a table is picked: lettered columns and blank cells."""
        self.sheet.set_sheet_data([[""] * 26 for _ in range(100)], reset_col_positions=True)
        self.sheet.headers([chr(65 + i) for i in range(26)])

    # sheet tabs ------------------------------------------------------------
    def new_tab(self, src: Tab | None = None) -> Tab:
        self._tab_seq += 1
        t = Tab(self._tab_seq)
        if src:  # same connection and database, empty grid
            t.server, t.uid, t.pwd, t.conn_name = src.server, src.uid, src.pwd, src.conn_name
            t.dbs, t._all_tables = list(src.dbs), list(src._all_tables)
            t.ui.update(db=src.ui["db"], top=src.ui["top"])
        prev, self.tab = self.tab, t
        self._new_sheet()
        self.tab = prev or t
        self.tabs.append(t)
        return t

    def add_tab(self):
        self._save_ui(self.shown)
        self.activate(self.new_tab(src=self.shown))
        self._save_tabs()

    def activate(self, t: Tab):
        if t is self.shown:
            return
        if self.shown:
            self._fx_flush()
            self._save_ui(self.shown)
            self.shown.sheet.pack_forget()
        self.tab = self.shown = t
        self._load_ui(t)
        t.sheet.pack(fill="both", expand=True)
        self._show_sql(t.sql_open)
        text, error = t.note
        self.l_msg.configure(text=text, style="Err.TLabel" if error else "Status.TLabel")
        self._progress()
        self._conn_icon()
        self._draw_tabs()
        self._refresh_buttons()
        self._marks()
        self._on_select()
        self._place_handle()
        if t.db.cn is None and not t.tried and t.server and (t.pwd or not t.uid):
            t.tried = True  # lazily: a restored tab connects when first opened
            self.connect()

    def cycle_tab(self, step: int):
        i = self.tabs.index(self.shown)
        self.activate(self.tabs[(i + step) % len(self.tabs)])

    def close_tab(self, t: Tab | None = None):
        t = t or self.shown
        if t._busy:
            return self.msg(f"{t.title} is still working; close it when done", error=True)
        if len(self.tabs) == 1:
            return self.msg("The last sheet cannot be closed", error=True)
        self._fx_flush()
        pending = self._pending(t)
        if pending and not messagebox.askyesno(
                "Close sheet", f"{t.title} has {pending}.\nDiscard them and close?", parent=self.root):
            return
        i = self.tabs.index(t)
        self.tabs.remove(t)
        if t is self.shown:
            t.sheet.pack_forget()
            self.shown = None
            self.activate(self.tabs[min(i, len(self.tabs) - 1)])
        t.sheet.destroy()
        t.db.close()
        self._draw_tabs()
        self._save_tabs()

    def _pending(self, t: Tab) -> str:
        """'2 edited / 1 new rows' for unsaved grid changes, '' when none."""
        prev, self.tab = self.tab, t
        try:
            if not self.cols:
                return ""
            new, edited = self._diff(self._rows())
        finally:
            self.tab = prev
        return f"{len(edited):,} edited / {len(new):,} new rows not pushed" if new or edited else ""

    def quit(self):
        self._fx_flush()
        pending = [f"{t.title}: {p}" for t in self.tabs if (p := self._pending(t))]
        if pending and not messagebox.askyesno(
                "Quit", "Unsaved changes:\n\n" + "\n".join(pending) + "\n\nDiscard them and quit?",
                parent=self.root):
            return
        self._save_tabs()
        for t in self.tabs:
            t.db.close()
        for job in (self._poll_id, self._mark_after):
            if job:
                self.root.after_cancel(job)
        self.root.destroy()

    def _in(self, t: Tab, fn):
        """Run fn with `t` as the current tab. For a tab in the background the ribbon
        fields are swapped to it for the call and swapped back after."""
        if t is self.shown:
            return fn()
        self._save_ui(self.shown)
        self.tab = t
        self._load_ui(t)
        try:
            fn()
        finally:
            self._save_ui(t)
            self.tab = self.shown
            self._load_ui(self.shown)

    def _save_ui(self, t: Tab | None):
        if t is None:
            return
        t.ui = {"db": self.v_db.get(), "table": self.v_table.get(), "where": self.v_where.get(),
                "top": self.v_top.get(), "sql": self.t_sql.get("1.0", "end-1c"), "dry": self.v_dry.get()}
        t.server, t.uid, t.pwd = self.v_server.get(), self.v_uid.get(), self.v_pwd.get()

    def _load_ui(self, t: Tab):
        def put(var, val):  # unchanged values are left alone, keeping the cursor
            if var.get() != val:
                var.set(val)
        for var, key in ((self.v_db, "db"), (self.v_table, "table"), (self.v_where, "where"),
                         (self.v_top, "top"), (self.v_dry, "dry")):
            put(var, t.ui[key])
        for var, val in ((self.v_server, t.server), (self.v_uid, t.uid), (self.v_pwd, t.pwd)):
            put(var, val)
        self.cb_db.configure(values=t.dbs)
        self.cb_table.configure(values=t._all_tables)
        if self.t_sql.get("1.0", "end-1c") != t.ui["sql"]:
            self.t_sql.delete("1.0", "end")
            self.t_sql.insert("1.0", t.ui["sql"])
            self.t_sql.edit_reset()

    def _save_tabs(self):
        self._save_ui(self.shown)
        self.settings["tabs"] = [
            {"conn": t.conn_name, "server": t.server, "user": t.uid, "db": t.ui["db"],
             "table": t.ui["table"], "where": t.ui["where"], "top": t.ui["top"], "sql": t.ui["sql"],
             "sql_open": t.sql_open} for t in self.tabs]
        self.settings["active_tab"] = self.tabs.index(self.shown) if self.shown in self.tabs else 0
        save_settings(self.settings)

    def _restore_tabs(self):
        env = (os.environ.get("DB_SERVER", ""), os.environ.get("DB_UID", ""),
               os.environ.get("DB_PWD", ""))
        saved = self.settings.get("tabs") or []
        if not saved:  # first start: last used saved connection, else .env
            last = conns.find(self.settings, self.settings.get("last_conn") or "")
            saved = [{"conn": last["name"], "server": last.get("server", ""),
                      "user": last.get("user", "")} if last else {"server": env[0], "user": env[1]}]
        for s in saved:
            t = self.new_tab()
            t.conn_name, t.server, t.uid = s.get("conn"), s.get("server", ""), s.get("user", "")
            c = conns.find(self.settings, t.conn_name) if t.conn_name else None
            if c and c.get("remember"):
                t.pwd = conns.get_password(c["name"])
            elif (t.server, t.uid) == env[:2]:
                t.pwd = env[2]
            for key in ("db", "table", "where", "top", "sql"):
                t.ui[key] = s.get(key) or t.ui[key]
            t.sql_open = bool(s.get("sql_open"))
        i = self.settings.get("active_tab", 0)
        self.activate(self.tabs[i if 0 <= i < len(self.tabs) else 0])

    def _draw_tabs(self):
        for w in self.tab_bar.winfo_children():
            w.destroy()
        for t in self.tabs:
            on = t is self.shown
            bg = C["surface"] if on else C["bg"]
            f = tk.Frame(self.tab_bar, bg=bg)
            f.pack(side="left", padx=(0, 1))
            row = tk.Frame(f, bg=bg)
            row.pack()
            t.label = tk.Label(row, text=self._tab_text(t), font=(FONT[0], 9, "bold" if on else "normal"),
                               bg=bg, fg=C["accent"] if on else C["text2"], padx=12, pady=3,
                               cursor="hand2")
            t.label.pack(side="left")
            x = tk.Label(row, text="×", font=FONT, bg=bg, fg=C["muted"], padx=2, cursor="hand2")
            x.pack(side="left", padx=(0, 6))
            tk.Frame(f, height=2, bg=C["accent"] if on else bg).pack(fill="x")
            t.label.bind("<Button-1>", lambda e, t=t: self.activate(t))
            t.label.bind("<Button-2>", lambda e, t=t: self.close_tab(t))
            x.bind("<Button-1>", lambda e, t=t: self.close_tab(t))
        plus = tk.Label(self.tab_bar, text="+", font=(FONT[0], 12), bg=C["bg"], fg=C["text2"],
                        padx=10, cursor="hand2")
        plus.pack(side="left")
        plus.bind("<Button-1>", lambda e: self.add_tab())

    @staticmethod
    def _tab_text(t: Tab) -> str:
        title = t.title
        return title if len(title) <= 30 else title[:29] + "…"

    def _conn_icon(self):
        self.i_conn.dot = "online" if self.shown.db.cn is not None else "offline"
        self.i_conn.redraw()

    def _show_sql(self, on: bool):
        self.v_sql_open.set(on)
        if on and not self.sql_box.winfo_manager():
            self.sql_box.pack(fill="x", before=self.grid_wrap)
        elif not on and self.sql_box.winfo_manager():
            self.sql_box.pack_forget()

    # formula bar / selection stats -----------------------------------------
    def _on_select(self):
        if self.tab is not self.shown:  # a background tab's job: the fx bar is not its own
            return
        self._fx_flush()
        self._fx_cell = None
        cur = self.sheet.get_currently_selected()
        if not cur or not self.cols or cur.column >= len(self.cols):
            self.v_name.set(""); self.v_fx.set(""); self.l_stats.configure(text="")
            return
        r, c = self._dr(cur.row), cur.column
        self.v_name.set(f"{self.cols[c].name}  {r + 1}")
        self.l_stats.configure(text=col_info(self.cols[c]))
        self.v_fx.set(self.fx_text(r, c))  # a formula cell shows its formula
        self._fx_cell = (self.tab, r, self.cols[c], self.v_fx.get())
        if (r, c) in self.bad:
            self.msg(f"{self.cols[c].name}: {self.bad[r, c]}", error=True)
        elif self.cols[c].picked:  # a key with a label: say what it means
            label = self.cols[c].key_label.get(self.v_fx.get())
            if label:
                self.msg(f"{self.cols[c].name} {self.v_fx.get()} = {label}")
        self.e_fx.state(["readonly"] if self.cols[c].readonly or self.cols[c].audit else ["!readonly"])
        self._sel_stats()

    def _sel_stats(self):
        cells = self.sheet.get_selected_cells()
        if len(cells) < 2 or len(cells) > 500_000:
            self.l_stats.configure(text="")
            return
        n = len(self.cols)
        count, nums = 0, []
        for r, c in cells:
            if c >= n:
                continue
            v = self.sheet.get_cell_data(self._dr(r), c)
            if v not in (None, ""):
                count += 1
                try:
                    nums.append(float(str(v).replace(",", "")))
                except ValueError:
                    pass
        parts = []
        if nums:
            total = sum(nums)
            parts += [f"Average: {total / len(nums):,.4g}", f"Count: {count:,}", f"Sum: {total:,.6g}"]
        elif count:
            parts.append(f"Count: {count:,}")
        self.l_stats.configure(text="     ".join(parts))

    def _fx_flush(self):
        """Write the fx bar into the cell it was showing, if it was changed there.
        Runs on Enter, on leaving the bar, on selecting another cell and before
        anything reads the grid, so a typed or pasted value is never dropped."""
        if not self._fx_cell:
            return
        tab, r, col, shown = self._fx_cell
        new = self.v_fx.get()
        if new == shown or tab is not self.tab or col not in self.cols:
            return
        if col.readonly or col.audit:
            self._fx_cell = (tab, r, col, new)
            return self.msg(f"{col.name} is read-only", error=True)
        if is_formula(new):
            self._retry_formula(r, self.cols.index(col))
        if col.big:
            self.tab.lk_hold = False
            col.nomatch.pop(new, None)
            col.missing.discard(new)
        new = resolve_choice(col, new)
        if new != self.v_fx.get():
            self.v_fx.set(new)
            self.msg(f"{col.name} {new} = {col.key_label.get(new, new)}")
        err = cell_error(col, new)
        if err:
            self.v_fx.set(shown)
            return self.msg(f"{col.name}: {err}", error=True)
        self._fx_cell = (tab, r, col, new)
        if r < self.sheet.get_total_rows():
            self.sheet.set_cell_data(r, self.cols.index(col), new, redraw=True)
            self._schedule_marks()

    def _begin_edit(self, event):
        """F2, Enter or double-click on a formula cell edits the formula, as in Excel."""
        if isinstance(event.key, str) and len(event.key) == 1:  # typing replaces the cell
            return event.value
        return self.fx_text(self._dr(event.row), event.column) if event.column < len(self.cols) else event.value

    def _validate_edit(self, event):
        """Typed cell edits must fit the column; pasted values are kept and marked
        instead (see _marks), so a big paste never loses data."""
        c = event.column
        if event.eventname == "end_edit_table" and c < len(self.cols):
            col = self.cols[c]
            value = "" if event.value is None else str(event.value)
            if is_formula(value):
                self._retry_formula(self._dr(event.row), c)
                return value
            value = resolve_choice(col, value)  # label -> key
            if col.big:
                self.tab.lk_hold = False
                col.nomatch.pop(value, None)  # typed again: ask again
                col.missing.discard(value)
                if value not in col.key_fields:  # a key or a name: the server says which
                    return value
            err = cell_error(col, value)
            if err:
                self.msg(f"{self.cols[c].name}: {err}", error=True)
                return None
            if col.choices is not None:
                if col.key_label.get(value):
                    self.msg(f"{col.name} {value} = {col.key_label[value]}")
                return value
        return event.value

    def _hover_cell(self, e):
        if not self.bad:
            return
        r, c = self.sheet.identify_row(e, allow_end=False), self.sheet.identify_column(e, allow_end=False)
        if r is not None and c is not None and (self._dr(r), c) in self.bad:
            self.msg(f"{self.cols[c].name}: {self.bad[self._dr(r), c]}", error=True)

    def _hover_header(self, e):
        c = self.sheet.identify_column(e, allow_end=False)
        if c is None or c >= len(self.cols):
            return
        self.l_stats.configure(text=col_info(self.cols[c]))

    def _fx_revert(self):
        if self._fx_cell:
            self.v_fx.set(self._fx_cell[3])

    def _fx_commit(self, _e=None):
        cur = self.sheet.get_currently_selected()
        if not cur or not self.cols or cur.column >= len(self.cols):
            return
        self._fx_flush()
        self.sheet.focus_set()
        if cur.row + 1 < self._shown_rows():
            self.sheet.select_cell(cur.row + 1, cur.column)
            self._on_select()

    # spreadsheet behaviour -----------------------------------------------------
    def _setup_excel(self):
        s, mt = self.sheet, self.sheet.MT
        self.filters: dict[str, dict] = {}  # column label -> excel_ops.matches spec
        self._drag: dict | None = None
        self._near_handle = False
        for label, fn, acc in (("Sort A to Z", lambda: self.sort(False), None),
                               ("Sort Z to A", lambda: self.sort(True), None),
                               ("Filter…", self.filter_dialog, "Ctrl+Shift+L"),
                               ("Clear filter", self.clear_filter, None),
                               ("Autofit column", self.autofit, None),
                               ("Lookup…", self.lookup_edit, None),
                               ("View lookup values…", self.lookup_view, None),
                               ("Label display…", self.label_display_menu, None),
                               ("Add scratch column", self.scratch_add, None),
                               ("Rename scratch column…", self.scratch_rename, None),
                               ("Remove scratch column", self.scratch_remove, None)):
            s.popup_menu_add_command(label, fn, table_menu=False, index_menu=False,
                                     empty_space_menu=False, accelerator=acc)
        for label, fn, acc in (("Fill down", self.fill_down, "Ctrl+D"),
                               ("Fill right", self.fill_right, "Ctrl+R"),
                               ("Filter by this value", self.filter_by_value, None),
                               ("Clear all filters", self.clear_all_filters, None)):
            s.popup_menu_add_command(label, fn, header_menu=False, index_menu=False,
                                     empty_space_menu=False, accelerator=acc)
        keys = {"<Control-d>": self.fill_down, "<Control-r>": self.fill_right,
                "<Control-y>": lambda: mt.redo(), "<Shift-space>": self.select_row,
                "<Control-L>": self.filter_dialog, "<Control-End>": lambda: self.jump_end()}
        for arrow, (dr, dc) in {"Up": (-1, 0), "Down": (1, 0), "Left": (0, -1), "Right": (0, 1)}.items():
            keys[f"<Control-{arrow}>"] = lambda dr=dr, dc=dc: self.jump(dr, dc, False)
            keys[f"<Control-Shift-{arrow}>"] = lambda dr=dr, dc=dc: self.jump(dr, dc, True)
        for seq, fn in keys.items():
            mt.bind(seq, lambda e, fn=fn: (fn(), "break")[1])

        # fill handle: a small square on the selection's bottom-right corner
        mt.create_rectangle(0, 0, 0, 0, tags=("aq_fill",), state="hidden", width=1)
        mt.create_rectangle(0, 0, 0, 0, tags=("aq_fill_prev",), state="hidden", dash=(4, 2), width=2)
        tag = "AQFillHandle"  # runs before tksheet's own bindings so a drag can be claimed
        mt.bindtags((tag, *mt.bindtags()))
        mt.bind_class(tag, "<Motion>", self._fh_motion)
        mt.bind_class(tag, "<ButtonPress-1>", lambda e: self._pt_click(e) or self._fh_press(e))
        mt.bind_class(tag, "<B1-Motion>", lambda e: self._pt_drag(e) or self._fh_drag(e))
        mt.bind_class(tag, "<ButtonRelease-1>", lambda e: self._pt_release(e) or self._fh_release(e))
        for seq in ("<Control-c>", "<Control-C>"):
            mt.bind_class(tag, seq, lambda e: self._fx_copy())
        for seq in ("<Control-v>", "<Control-V>"):
            mt.bind_class(tag, seq, lambda e: self._fx_paste())
        self._pt = None  # pointing at cells while typing a formula
        keys = mt.text_editor_user_bound_keys
        for arrow, d in {"<Up>": (-1, 0), "<Down>": (1, 0), "<Left>": (0, -1), "<Right>": (0, 1)}.items():
            keys[arrow] = lambda e, d=d: self._pt_arrow(*d)
        keys["<Key>"] = lambda e: self._pt_end()
        s.bind("<<SheetRedrawn>>", lambda e: self._place_handle())

    # formulas: copy / paste ---------------------------------------------------
    def _fx_copy(self):
        """Ctrl+C: other programs get values; a paste here gets the formulas back."""
        cur = self.sheet.get_currently_selected()
        self._clip = None
        if not cur or not self.tab.fx:
            return
        b = cur.box
        found = {(r - b.from_r, c - b.from_c): (self._dr(r), c, e["f"])
                 for r in range(b.from_r, b.upto_r) for c in range(b.from_c, min(b.upto_c, len(self.cols)))
                 if (e := self._fx_get(self._dr(r), c))}
        if found:
            self.root.after_idle(lambda: setattr(self, "_clip", (self._clipboard(), found)))

    def _clipboard(self) -> str:
        try:
            return self.root.clipboard_get()
        except tk.TclError:
            return ""

    def _fx_paste(self):
        clip = getattr(self, "_clip", None)
        if not clip or clip[0] != self._clipboard():
            return

        def after():  # tksheet pasted the values; put the formulas back, moved
            cur = self.sheet.get_currently_selected()
            if not cur:
                return
            b = cur.box
            for (dr, dc), (src_r, src_c, f) in clip[1].items():
                r, c = b.from_r + dr, b.from_c + dc
                if r < self._shown_rows() and c < len(self.cols) and self._writable(c):
                    gi = self._dr(r)
                    self.sheet.set_cell_data(gi, c, fx_shift(f, gi - src_r, c - src_c, self.cols), redraw=False)
            self.sheet.redraw()
            self._schedule_marks()
            self._on_select()

        self.root.after_idle(after)

    # formulas: point at cells while typing one, as in Excel -------------------
    def _pt_widget(self):
        """The widget a formula is being typed in: the cell editor or the fx bar."""
        te = self.sheet.MT.text_editor
        if te.open and te.window and te.tktext.get("1.0", "end-1c").startswith("="):
            return te.tktext
        try:
            if self.root.focus_get() is self.e_fx and self.v_fx.get().startswith("="):
                return self.e_fx
        except (KeyError, tk.TclError):
            pass
        return None

    def _pt_ready(self, w) -> bool:
        """At a spot where a reference fits: after =, (, comma or an operator."""
        if self._pt and self._pt["w"] is w:
            return True
        if isinstance(w, tk.Text):
            before = w.get("1.0", "insert")
        else:
            before = w.get()[:w.index("insert")]
        return before.rstrip()[-1:] in tuple("=(,+-*/<>&%^|")

    def _pt_ref(self, r: int, c: int, r2: int | None = None, c2: int | None = None) -> str:
        """Reference text for displayed cell r, c (or a range to r2, c2)."""
        gi = self._dr(r)
        if r2 is not None and (r2, c2) != (r, c):
            a, b = sorted((gi, self._dr(r2))), sorted((c, c2))
            return f"{num2alpha(b[0])}{a[0] + 1}:{num2alpha(b[1])}{a[1] + 1}"
        own = self._pt_own_row()
        if self.settings.get("formula_refs", "names") == "names" and gi == own:
            return fx_ref_name(self.cols[c])
        return f"{num2alpha(c)}{gi + 1}"

    def _pt_own_row(self) -> int | None:
        te = self.sheet.MT.text_editor
        if te.open:
            return self._dr(te.coords[0])
        return self._fx_cell[1] if self._fx_cell else None

    def _pt_put(self, w, text: str):
        pt = self._pt if self._pt and self._pt["w"] is w else None
        if isinstance(w, tk.Text):
            if pt:
                w.delete(pt["a"], pt["b"])
                w.mark_set("insert", pt["a"])
            a = w.index("insert")
            w.insert("insert", text)
            b = w.index("insert")
        else:
            if pt:
                w.delete(pt["a"], pt["b"])
                w.icursor(pt["a"])
            a = w.index("insert")
            w.insert("insert", text)
            b = w.index("insert")
        w.focus_set()
        return a, b

    def _pt_click(self, e):
        w = self._pt_widget()
        if w is None or not self._pt_ready(w):
            self._pt = None
            return None
        r = self.sheet.identify_row(e, allow_end=False)
        c = self.sheet.identify_column(e, allow_end=False)
        if r is None or c is None or c >= len(self.cols):
            return None
        a, b = self._pt_put(w, self._pt_ref(r, c))
        self._pt = {"w": w, "a": a, "b": b, "r": r, "c": c, "drag": True}
        return "break"

    def _pt_drag(self, e):
        pt = self._pt
        if not pt or not pt.get("drag"):
            return None
        r = self.sheet.identify_row(e, allow_end=False)
        c = self.sheet.identify_column(e, allow_end=False)
        if r is not None and c is not None and c < len(self.cols):
            a, b = self._pt_put(pt["w"], self._pt_ref(pt["r"], pt["c"], r, c))
            pt.update(a=a, b=b)
        return "break"

    def _pt_release(self, e):
        if self._pt and self._pt.get("drag"):
            self._pt["drag"] = False
            return "break"
        return None

    def _pt_arrow(self, dr: int, dc: int):
        """Arrows right after an operator point at a cell instead of moving the caret."""
        te = self.sheet.MT.text_editor
        w = self._pt_widget()
        if w is None or w is not te.tktext or not self._pt_ready(w):
            self._pt = None
            return None
        r, c = (self._pt["r"], self._pt["c"]) if self._pt else te.coords
        r = max(0, min(self._shown_rows() - 1, r + dr))
        c = max(0, min(len(self.cols) - 1, c + dc))
        a, b = self._pt_put(w, self._pt_ref(r, c))
        self._pt = {"w": w, "a": a, "b": b, "r": r, "c": c, "drag": False}
        self.sheet.see(r, c, redraw=True)
        return "break"

    def _pt_end(self):
        self._pt = None  # typing ends pointing; the next operator starts it again

    def _dr(self, r: int) -> int:  # displayed row -> data row (differs while filtered)
        return self.sheet.displayed_row_to_data(r)

    def _shown_rows(self) -> int:
        return len(self.sheet.MT.row_positions) - 1

    def _cell(self, r: int, c: int) -> str:
        v = self.sheet.get_cell_data(self._dr(r), c)
        return "" if v is None else str(v)

    def _writable(self, c: int) -> bool:
        col = self.cols[c]
        return not (col.readonly or col.audit)

    # column moves: tksheet moves data + headers; keep self.cols / self.orig in step
    def _sync_columns(self):
        n = len(self.cols)
        hdr = list(self.sheet.headers())[:n]
        labels = [c.label for c in self.cols]
        if hdr == labels or sorted(hdr) != sorted(labels):
            return
        perm = [labels.index(h) for h in hdr]
        self.cols = [self.cols[p] for p in perm]
        for rid, o in self.orig.items():
            self.orig[rid] = [o[p] for p in perm]

    # sort / filter
    def _col(self) -> int | None:
        cur = self.sheet.get_currently_selected()
        return cur.column if cur and self.cols and cur.column < len(self.cols) else None

    def sort(self, reverse: bool):
        c = self._col()
        if c is None:
            return
        self.sheet.sort_rows_by_column(c, reverse=reverse)
        self._apply_filters()
        self._schedule_marks()

    def autofit(self):
        c = self._col()
        if c is not None:
            longest = max([len(self.cols[c].label)] + [len(self._cell(r, c)) for r in range(min(500, self._shown_rows()))])
            ch = tkfont.Font(family=CELL[0], size=CELL[1]).measure("0")
            self.sheet.column_width(c, min(600, max(60, longest * ch + 20)))
            self.sheet.redraw()

    def filter_dialog(self):
        c = self._col()
        if c is not None:
            FilterDialog(self, c)

    def filter_by_value(self):
        cur = self.sheet.get_currently_selected()
        if not cur or cur.column >= len(self.cols):
            return
        v = self._cell(cur.row, cur.column)
        self.set_filter(cur.column, {"values": {v or excel_ops.BLANK}, "text": ""})

    def set_filter(self, c: int, spec: dict | None):
        label = self.cols[c].label
        if spec is None:
            self.filters.pop(label, None)
        else:
            self.filters[label] = spec
        self._apply_filters()

    def clear_filter(self):
        c = self._col()
        if c is not None:
            self.set_filter(c, None)

    def clear_all_filters(self):
        self.filters = {}
        self._apply_filters()

    def _apply_filters(self):
        idx = {c.label: i for i, c in enumerate(self.cols)}
        self.filters = {k: v for k, v in self.filters.items() if k in idx}
        if not self.filters:
            self.sheet.display_rows("all", all_rows_displayed=True, redraw=True)
        else:
            # blank spare rows under the data are not part of it: never filtered
            keep = [r for r, row in enumerate(self._rows())
                    if self._spare(row) or all(excel_ops.matches(row[idx[k]], spec)
                                               for k, spec in self.filters.items())]
            self.sheet.display_rows(keep, all_rows_displayed=False, redraw=True)
        self._marks()

    # fill down / right / handle
    def fill_down(self):
        cur = self.sheet.get_currently_selected()
        if not cur or not self.cols:
            return
        b = cur.box
        if b.upto_r - b.from_r > 1:   # first row of the selection into the rest
            self._fill((b.from_r, b.from_c, b.from_r + 1, b.upto_c), end_r=b.upto_r, copy=True)
        elif b.from_r > 0:            # single row: copy the row above
            self._fill((b.from_r - 1, b.from_c, b.from_r, b.upto_c), end_r=b.from_r + 1, copy=True)

    def fill_right(self):
        cur = self.sheet.get_currently_selected()
        if not cur or not self.cols:
            return
        b = cur.box
        if b.upto_c - b.from_c > 1:
            self._fill((b.from_r, b.from_c, b.upto_r, b.from_c + 1), end_c=b.upto_c, copy=True)
        elif b.from_c > 0:
            self._fill((b.from_r, b.from_c - 1, b.upto_r, b.from_c), end_c=b.from_c + 1, copy=True)

    def _fill(self, box, end_r: int | None = None, end_c: int | None = None, copy: bool = False):
        """Extend the source `box` (displayed from_r, from_c, upto_r, upto_c) down to
        row end_r or right to column end_c, as a series or a plain copy."""
        r0, c0, r1, c1 = box
        if end_r is not None and end_r > self._shown_rows():
            if self.filters:
                end_r = self._shown_rows()
                self.msg("Clear filters to fill past the last row", error=True)
            else:
                self.sheet.insert_rows(end_r - self._shown_rows(), redraw=False)
        writes: dict[tuple[int, int], str] = {}
        if end_r is not None:
            for c in range(c0, c1):
                if not self._writable(c):
                    continue
                src = [self.fx_text(self._dr(r), c) for r in range(r0, r1)]
                if any(is_formula(v) for v in src):  # formulas copy, their references move down
                    vals = [fx_shift(src[i % len(src)], r1 + i - (r0 + i % len(src)), 0, self.cols)
                            for i in range(end_r - r1)]
                else:
                    vals = ([src[i % len(src)] for i in range(end_r - r1)] if copy
                            else excel_ops.fill_series(src, end_r - r1))
                for i, v in enumerate(vals):
                    writes[(r1 + i, c)] = v
        elif end_c is not None:
            for r in range(r0, r1):
                src = [self.fx_text(self._dr(r), c) for c in range(c0, c1)]
                if any(is_formula(v) for v in src):
                    vals = [fx_shift(src[i % len(src)], 0, c1 + i - (c0 + i % len(src)), self.cols)
                            for i in range(end_c - c1)]
                else:
                    vals = ([src[i % len(src)] for i in range(end_c - c1)] if copy
                            else excel_ops.fill_series(src, end_c - c1))
                for i, v in enumerate(vals):
                    if self._writable(c1 + i):
                        writes[(r, c1 + i)] = v
        if not writes:
            return
        rows = sorted({r for r, _ in writes})
        cols = sorted({c for _, c in writes})
        if not self.filters:  # one block write = one undo step
            ra, rb, ca, cb = rows[0], rows[-1] + 1, cols[0], cols[-1] + 1
            block = [[writes.get((r, c), self._cell(r, c)) for c in range(ca, cb)] for r in range(ra, rb)]
            self.sheet.set_data(self.sheet.span((ra, ca, rb, cb)), data=block, undo=True, redraw=True)
        else:
            for (r, c), v in writes.items():
                self.sheet.set_cell_data(self._dr(r), c, v, redraw=False)
            self.sheet.redraw()
        self._schedule_marks()
        self._on_select()

    def _handle_xy(self):
        cur = self.sheet.get_currently_selected()
        mt = self.sheet.MT
        if not cur or not self.cols or not self._shown_rows():
            return None
        b = cur.box
        try:
            return mt.col_positions[b.upto_c], mt.row_positions[b.upto_r], b
        except IndexError:
            return None

    def _place_handle(self):
        mt, pos = self.sheet.MT, self._handle_xy()
        if pos is None:
            mt.itemconfigure("aq_fill", state="hidden")
            return
        x, y, _ = pos
        mt.coords("aq_fill", x - 3, y - 3, x + 3, y + 3)
        mt.itemconfigure("aq_fill", state="normal", fill=C["accent"], outline=C["surface"])
        mt.tag_raise("aq_fill")

    def _fh_motion(self, e):
        mt, pos = self.sheet.MT, self._handle_xy()
        near = bool(pos) and abs(mt.canvasx(e.x) - pos[0]) <= 5 and abs(mt.canvasy(e.y) - pos[1]) <= 5
        if near != self._near_handle:
            mt.configure(cursor="crosshair" if near else "")
            self._near_handle = near
        return "break" if near else None

    def _fh_press(self, e):
        if not self._near_handle:
            return None
        pos = self._handle_xy()
        self._drag = {"box": pos[2], "copy": bool(e.state & 0x4), "end_r": None, "end_c": None}
        return "break"

    def _fh_drag(self, e):
        if not self._drag:
            return None
        mt, b = self.sheet.MT, self._drag["box"]
        x, y = mt.canvasx(e.x), mt.canvasy(e.y)
        rp, cp = mt.row_positions, mt.col_positions
        r = bisect.bisect_right(rp, y) - 1
        if y >= rp[-1]:  # past the last row: grow by row height
            r = len(rp) - 2 + int((y - rp[-1]) // max(1, rp[1] - rp[0])) + 1
        c = min(max(bisect.bisect_right(cp, x) - 1, 0), len(cp) - 2)
        down, right = r - (b.upto_r - 1), c - (b.upto_c - 1)
        self._drag["end_r"] = self._drag["end_c"] = None
        if down > 0 and down >= right:
            self._drag["end_r"] = r + 1
            row_h = rp[1] - rp[0]
            bottom = rp[r + 1] if r + 1 < len(rp) else rp[-1] + (r + 2 - len(rp)) * row_h
            x0, y0, x1, y1 = cp[b.from_c], rp[b.from_r], cp[b.upto_c], bottom
        elif right > 0:
            self._drag["end_c"] = c + 1
            x0, y0, x1, y1 = cp[b.from_c], rp[b.from_r], cp[c + 1], rp[b.upto_r]
        else:
            mt.itemconfigure("aq_fill_prev", state="hidden")
            return "break"
        mt.coords("aq_fill_prev", x0, y0, x1, y1)
        mt.itemconfigure("aq_fill_prev", state="normal", outline=C["accent"])
        mt.tag_raise("aq_fill_prev")
        if e.y > mt.winfo_height() - 12 and r < self._shown_rows():
            self.sheet.see(min(r + 1, self._shown_rows() - 1), c)
        return "break"

    def _fh_release(self, e):
        if not self._drag:
            return None
        d, self._drag = self._drag, None
        self.sheet.MT.itemconfigure("aq_fill_prev", state="hidden")
        b = d["box"]
        if d["end_r"] or d["end_c"]:
            self._fill((b.from_r, b.from_c, b.upto_r, b.upto_c), end_r=d["end_r"], end_c=d["end_c"],
                       copy=d["copy"])
        return "break"

    # keyboard navigation
    def select_row(self):
        cur = self.sheet.get_currently_selected()
        if cur:
            self.sheet.select_row(cur.row)

    def jump(self, dr: int, dc: int, extend: bool):
        """Ctrl+Arrow: to the edge of the current block of values, as in Excel."""
        cur = self.sheet.get_currently_selected()
        if not cur or not self.cols:
            return
        nr, nc = self._shown_rows(), len(self.cols)
        r, c = cur.row, cur.column

        def filled(r, c):
            return 0 <= r < nr and 0 <= c < nc and self._cell(r, c) != ""

        if filled(r, c) and filled(r + dr, c + dc):
            while filled(r + dr, c + dc):
                r, c = r + dr, c + dc
        else:
            r, c = r + dr, c + dc
            while 0 <= r < nr and 0 <= c < nc and not filled(r, c):
                r, c = r + dr, c + dc
        r, c = min(max(r, 0), nr - 1), min(max(c, 0), nc - 1)
        self._go_to(cur, r, c, extend)

    def jump_end(self):
        cur = self.sheet.get_currently_selected()
        if self.cols and self._shown_rows():
            self._go_to(cur, self._shown_rows() - 1, len(self.cols) - 1, False)

    def _go_to(self, cur, r: int, c: int, extend: bool):
        if extend and cur:
            ar, ac = cur.row, cur.column
            self.sheet.create_selection_box(min(ar, r), min(ac, c), max(ar, r) + 1, max(ac, c) + 1)
            self.sheet.set_currently_selected(ar, ac)
        else:
            self.sheet.select_cell(r, c)
        self.sheet.see(r, c)
        self._on_select()

    # background jobs -------------------------------------------------------
    def bg(self, label: str, work, done=None, cancel=False):
        """Run work() in a thread, then done(result, seconds) on the UI thread with
        the same tab current. work must not read tab state (self.db etc.):
        capture it before, the user may switch tabs meanwhile. `cancel` offers the
        Cancel button (for reads: Db.cancel stops the running statement)."""
        tab = self.tab
        t0 = dt.datetime.now()
        token = object()  # identifies this job; a cancelled job's result is dropped
        tab.job = (label, t0, cancel, token)
        self._running.add(token)
        tab.db.begin()
        self._busy = True
        self._refresh_buttons()
        self.msg(label + " …")

        def run():
            try:
                res = work()
                self.jobs.put((tab, done, res, None, t0, token))
            except Exception as e:  # noqa: BLE001 - surfaced in UI
                self.jobs.put((tab, done, None, e, t0, token))

        threading.Thread(target=run, daemon=True).start()

    def _poll(self):
        try:
            while True:
                tab, done, res, err, t0, token = self.jobs.get_nowait()
                self._running.discard(token)
                old = self._abandoned.pop(token, None)
                if old is not None:  # its connection was replaced after a cancel
                    threading.Thread(target=old.close, daemon=True).start()
                if tab not in self.tabs or not tab.job or tab.job[3] is not token:
                    continue  # cancelled (or tab closed): nothing to show
                secs = (dt.datetime.now() - t0).total_seconds()

                def finish():
                    self._busy = False
                    self.tab.job = None
                    if err:
                        self._show_error(err)
                    elif done:
                        done(res, secs)
                        if tab.db.notes:  # PRINT / trigger messages of that job
                            text, error = tab.note
                            self.msg(f"{text}\nServer: " + "\n".join(tab.db.notes), error)
                    self._refresh_buttons()
                    self._marks()

                self._in(tab, finish)
                if tab is not self.shown:  # ribbon buttons belong to the shown tab again
                    self._refresh_buttons()
                    self._marks()
                self._conn_icon()
        except queue.Empty:
            pass
        self._progress()
        self._poll_id = self.root.after(50, self._poll)

    def _progress(self):
        """Running job of the shown tab: elapsed time and rows read so far."""
        t = self.shown
        if not t or not t.job:
            return
        label, t0 = t.job[:2]
        secs = (dt.datetime.now() - t0).total_seconds()
        if secs < 0.5:
            return
        n = t.db.fetched if t.db.active is not None else 0
        rows = f" · {n:,} rows" if n else ""
        self.l_msg.configure(text=f"{label} …{rows} · {secs:.0f}s", style="Status.TLabel")

    def _refresh_buttons(self):
        for b in self.busy_buttons:
            b.state(["disabled"] if self._busy else ["!disabled"])
        if not self._busy:
            self._set_table_state(bool(self.cols))
        job = self.shown.job if self.shown else None
        if job and job[2]:
            self.b_cancel.pack(side="right", padx=(8, 0))
        else:
            self.b_cancel.pack_forget()

    def cancel_job(self):
        """Cancel always frees the sheet at once. The server is asked to stop;
        if the connection has not answered within 2 s (blocked server, dead
        network) the sheet gets a fresh connection and the old one is closed
        when its statement finally returns."""
        t = self.shown
        if not t or not t.job:
            return
        token, db = t.job[3], t.db
        db.cancel()
        t.job, t._busy = None, False
        self._refresh_buttons()
        self._marks()
        self.msg("Cancelled")

        def check():
            if token not in self._running or t not in self.tabs or t.db is not db:
                return  # it stopped in time
            self._abandoned[token] = db  # closed when its worker returns
            fresh = Db()
            t.db = fresh
            if db.args:
                def done(_, secs):
                    self.msg("Cancelled · reconnected")
                self._in(t, lambda: self.bg("Reconnecting", lambda: fresh.connect(*db.args), done))

        self.root.after(2000, check)

    def _show_error(self, err):
        if isinstance(err, RowError):
            where = f"Grid row {err.grid_row + 1}: " if err.grid_row is not None else ""
            self.msg(f"{where}{err.msg}  Nothing committed.", error=True)
            if err.grid_row is not None:
                self.clear_all_filters()
                self.sheet.see(err.grid_row, 0)
                self.sheet.select_row(err.grid_row)
        else:
            self.msg(clean_err(err), error=True)

    def msg(self, text, error=False):
        """Bottom-bar message. Multi-line text shows on one line; click the bar
        for the full text of recent messages (MessageLog)."""
        text = str(text)
        if error or "\n" in text or len(text) > 160:
            self.log.append((dt.datetime.now().strftime("%H:%M:%S"), self.tab.title, text, error))
            del self.log[:-200]
        line = " · ".join(ln.strip() for ln in text.splitlines() if ln.strip())
        self.tab.note = (line, error)
        if self.tab is self.shown:
            self.l_msg.configure(text=line, style="Err.TLabel" if error else "Status.TLabel")

    def _uncaught(self, exc, val, tb):
        """Errors in UI callbacks: bottom bar + log, never a silent stderr line."""
        import traceback
        detail = "".join(traceback.format_exception(exc, val, tb))
        try:
            self.msg(f"Unexpected error: {val!r}\n\n{detail}", error=True)
        except Exception:  # noqa: BLE001 - while starting up / shutting down
            sys.stderr.write(detail)

    def _set_table_state(self, on: bool):
        for b in (self.b_insert, self.b_update, self.b_paste, self.b_import,
                  self.b_export):
            b.state(["!disabled"] if on else ["disabled"])

    # connection / source ---------------------------------------------------
    def connect(self):
        server, uid, pwd = self.v_server.get().strip(), self.v_uid.get().strip(), self.v_pwd.get()
        db = self.db
        self.conn_text = "Not connected"
        self.tab.server = server

        def work():
            db.connect(server, uid, pwd)
            return db.databases()

        def done(dbs, secs):
            self.tab.dbs = dbs
            self.cb_db.configure(values=dbs)
            self.conn_text = f"Connected · {uid or 'Windows auth'} @ {server}"
            if self.v_db.get() in dbs:
                self.pick_db()
            self.msg(f"Connected · {len(dbs)} databases · {secs:.1f}s")

        self.bg("Connecting", work, done)

    def pick_db(self):
        database = self.v_db.get()

        def done(tables, secs):
            self._all_tables = tables
            self.cb_table.configure(values=tables)
            self.msg(f"{database} · {len(tables)} tables")

        db = self.db
        self.bg("Listing tables", lambda: db.tables(database), done)

    def _lookup_prefs(self) -> dict:
        """Lookups chosen from the header menu, copied for a worker thread."""
        return {"lookups": dict(self.settings.get("lookups", {})),
                "fk_labels": dict(self.settings.get("fk_labels", {}))}

    def _lookup_spec(self) -> dict | None:
        """The lookup file from Settings, read fresh so edits apply on the next Load."""
        try:
            return load_lookup_file(self.settings.get("lookup_file", ""))
        except (OSError, ValueError) as e:
            self.msg(f"Lookup file: {e}", error=True)
            return None

    def _paste_where(self, e):
        """A pasted column - header line, then one value per line, as copied from a
        spreadsheet - becomes `header IN (values)`. Anything else pastes as text."""
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            return "break"
        clause = in_clause(text, {c.name.lower(): c for c in self.table_cols or self.cols})
        if clause is None:
            return None  # plain paste (class binding)
        w = e.widget
        try:
            w.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        before = w.get().rstrip()
        w.delete(0, "end")
        w.insert(0, f"{before} AND {clause}" if before else clause)
        return "break"

    def _filter_tables(self, e):
        if e.keysym in ("Return", "Up", "Down", "Escape"):
            return
        t = self.v_table.get().lower()
        self.cb_table.configure(values=[x for x in self._all_tables if t in x.lower()][:500])

    def _source(self) -> tuple[str, str] | None:
        database, table = self.v_db.get(), self.v_table.get().strip()
        if self.db.cn is None:
            self.msg("Connect first", error=True)
        elif not database or not table:
            self.msg("Pick a database and a table", error=True)
        else:
            return database, table
        return None

    def load(self):
        src = self._source()
        if not src:
            return
        database, table = src
        where = self.v_where.get()
        if ";" in where:
            return self.msg("Where clause cannot contain ';'", error=True)
        try:
            top = max(1, int(self.v_top.get()))
        except ValueError:
            return self.msg("Top must be a number", error=True)

        db = self.db
        spec = self._lookup_spec()
        if spec is None:
            return
        prefs = self._lookup_prefs()

        def work():
            cols = db.columns(database, table)
            db.lookups(database, table, cols, spec, prefs)
            return database, table, cols, db.load(database, table, cols, where, top)

        def done(res, secs):
            database, table, cols, rows = res
            self._set_grid(database, table, cols, rows)
            self.msg(f"Loaded {len(rows):,} rows · {secs:.1f}s")

        self.bg("Loading", work, done, cancel=True)

    def blank(self):
        src = self._source()
        if not src:
            return
        database, table = src

        def done(res, secs):
            database, table, cols = res
            self._set_grid(database, table, cols, [])
            self.msg("Paste rows from Excel into the grid, or use Paste with headers to map "
                     "columns by name")

        db = self.db
        spec = self._lookup_spec()
        if spec is None:
            return
        prefs = self._lookup_prefs()

        def work():
            cols = db.columns(database, table)
            db.lookups(database, table, cols, spec, prefs)
            return database, table, cols

        self.bg("Reading columns", work, done)

    def toggle_sql(self):
        self.tab.sql_open = not self.tab.sql_open
        self._show_sql(self.tab.sql_open)
        if not self.tab.sql_open:
            return
        if not self.t_sql.get("1.0", "end").strip() and self.v_table.get().strip():
            self.t_sql.insert("1.0", f"SELECT *\nFROM {self.v_table.get().strip()}\nWHERE 1 = 1")
        self.t_sql.focus_set()

    def run_query(self):
        if self.db.cn is None:
            return self.msg("Connect first", error=True)
        database, table = self.v_db.get(), self.v_table.get().strip()
        sql = self.t_sql.get("1.0", "end").strip()
        err = check_select(sql)
        if err:
            return self.msg(err, error=True)
        spec = self._lookup_spec()
        if spec is None:
            return
        prefs = self._lookup_prefs()
        try:
            top = max(1, int(self.v_top.get()))
        except ValueError:
            return self.msg("Top must be a number", error=True)

        db = self.db

        def work():
            names, rows = db.query(database, sql, top)
            srcs = db.sources(database, sql)
            tdb, target, table_cols = edit_target(db, database, table, srcs)
            if table_cols:
                db.lookups(tdb, target, table_cols, spec, prefs)
            tables = db.tables(tdb) if tdb and tdb != database else None
            return tdb or database, target, table_cols, names, rows, srcs, tables

        def done(res, secs):
            tdb, target, table_cols, names, rows, srcs, tables = res
            if tables is not None:  # the result lives in another database: follow it
                self.v_db.set(tdb)
                self._all_tables = tables
                self.cb_table.configure(values=tables)
            database = tdb
            if target != table:
                self.v_table.set(target)
            cols, note = self._map_query_cols(database, target, table_cols, names, srcs,
                                              table_refs(sql, target) if target else 0)
            self._set_grid(database, target, cols, rows, table_cols, from_query=True)
            capped = f" (Top {top:,} reached)" if len(rows) == top else ""
            self.msg(f"Query returned {len(rows):,} rows{capped} · {secs:.1f}s{note}")

        self.bg("Running query", work, done, cancel=True)

    def _map_query_cols(self, database, table, table_cols, names, srcs, refs: int) -> tuple[list[Col], str]:
        """Editable only when SQL Server says the column comes from `table` itself.
        A column listed twice is the same column when the table is referenced once
        (`SELECT IssueDate, * FROM t`): both copies stay editable and are kept in
        step (see _sync_twins). With a self-join (refs > 1) the copies may be
        different rows, so they are read-only."""
        def shown(nm):
            return Col(nm, "", False, False, False, external=True)

        if not table:
            return [shown(nm) for nm in names], ""
        if srcs is None or len(srcs) != len(names):
            return [shown(nm) for nm in names], (" · column sources could not be resolved, "
                                                 "result is read-only")
        want = (database.lower(), *(x.lower() for x in split_table(table)))
        by_name = {c.name.lower(): c for c in table_cols}
        hit = [by_name.get(s[3].lower()) if s and tuple(x.lower() for x in s[:3]) == want
               else None for s in srcs]
        counts: dict[str, int] = {}
        for c in hit:
            if c:
                counts[c.name] = counts.get(c.name, 0) + 1
        cols, dup = [], set()
        for nm, c in zip(names, hit):
            if c and (counts[c.name] == 1 or refs == 1):
                c = copy.copy(c)
                c.label = nm if nm.lower() == c.name.lower() else f"{nm} ({c.name})"
                cols.append(c)
            else:
                if c:
                    dup.add(c.name)
                cols.append(shown(nm))
        ext = sum(c.external and c.name not in dup for c in cols)
        note = f" · {ext} read-only column{'s' if ext != 1 else ''} not from {table}" if ext else ""
        if dup:
            note += f" · {', '.join(sorted(dup))} appear more than once, read-only"
        return cols, note

    def _set_grid(self, database, table, cols, rows, table_cols=None, from_query=False):
        self.database, self.table, self.cols = database, table, cols
        self.table_cols = cols if table_cols is None else table_cols
        self.from_query = from_query
        self.keys = [c for c in cols if c.is_pk and not c.external]
        self.orig = {}
        data = []
        for i, r in enumerate(rows):
            rid = str(i)
            self.orig[rid] = r
            data.append(r + [rid])
        data += [[""] * (len(cols) + 1) for _ in range(SPARE_ROWS)]
        n = len(cols)
        seen: dict[str, int] = {}
        for c in cols:  # header text must be unique (column moves are tracked by it);
            k = c.label.lower()  # a repeat gets invisible zero-width spaces appended
            c.label += "\u200b" * seen.get(k, 0)
            seen[k] = seen.get(k, 0) + 1
        self._show(cols, data, self.orig)

    def _show(self, cols, data, orig):
        """Put columns, rows (with row ids) and loaded values on the sheet, adding
        label columns for lookups shown in a separate column."""
        cols, data, orig = self._add_label_cols(cols, data, orig)
        self.cols, self.orig = cols, orig
        n = len(cols)
        self.filters = {}
        self.sheet.display_rows("all", all_rows_displayed=True)
        self.sheet.disp = {}  # old display rules may name columns that changed
        self.sheet.set_sheet_data(data, reset_col_positions=True)
        self.sheet.headers([c.label for c in cols] + ["__rid"])
        self.sheet.display_columns(list(range(n)), all_columns_displayed=False)
        self.sheet.readonly_columns([i for i, c in enumerate(cols) if c.readonly or c.audit], True)
        self._dropdowns()
        self._fit_columns(data)
        self._marks()

    def _label_mode(self, col: "Col") -> str:
        """same / value / separate / off for a lookup column: its own choice, else Settings."""
        own = self.settings.get("label_modes", {}).get(col_key(self.database, self.table, col.name)) \
            if self.table else None
        return own or self.settings.get("label_display", "same")

    def _label_fmt(self, mode: str) -> str:
        return "{label}" if mode == "value" else self.settings.get("label_format", "{id} · {label}")

    def _sub_formats(self):
        """How each sub lookup's values show: its own mode ("separate" = id · value)."""
        def walk(c):
            for sub in c.subs.values():
                m = self._label_mode(sub)
                sub.fmt = "" if m == "off" else self._label_fmt(m)
                walk(sub)
        for c in self.cols:
            if c.of is None:
                walk(c)

    def _add_label_cols(self, cols, data, orig):
        """A read-only "<column> ▸ label" column after each lookup column shown in
        separate mode. Never written; refreshed from the id as it changes."""
        mark = self.settings.get("mismatch", "both") in ("mark", "both")
        self._sub_formats()
        out, extra = [], {}
        for i, c in enumerate(cols):
            out.append(c)
            if c.of is None and c.labelled and self._label_mode(c) == "separate":
                for part, field in enumerate(c.label_fields):
                    v = Col(f"{c.name} ▸ {field}", "", False, False, False, external=True)
                    v.of, v.part, v.label = c, part, f"{c.label} ▸ {field}"
                    out.append(v)
                extra[i] = c
        if not extra:
            return cols, data, orig
        n = len(cols)

        def widen(r):
            row = []
            for i in range(n):
                row.append(r[i])
                if i in extra:
                    c = extra[i]
                    row += [label_text(c, r[i] or "", mark, p) for p in range(len(c.label_fields))]
            return row

        return out, [widen(r) + r[n:] for r in data], {k: widen(o) for k, o in orig.items()}

    def _relabel(self, add: tuple[int, "Col"] | None = None, drop: "Col | None" = None):
        """Rebuild the columns after a display or lookup change, or to add (before
        index) / drop a scratch column; edits and formulas stay."""
        if not self.cols:
            return
        n = len(self.cols)
        rows = self._rows()
        keep = [i for i, c in enumerate(self.cols) if c.of is None and c is not drop]
        if add:
            at_col = next((k for k, i in enumerate(keep) if i >= add[0]), len(keep))
            keep.insert(at_col, None)  # None: the new column, empty
        cols = [self.cols[i] if i is not None else add[1] for i in keep]
        data = [[r[i] if i is not None else "" for i in keep] + [r[n]] for r in rows]
        orig = {k: [o[i] if i is not None else "" for i in keep] for k, o in self.orig.items()}
        old_data, old_cols = self.sheet.MT.data, list(self.cols)
        formulas = [(gi, old_cols[ci], e) for gi, row in enumerate(old_data)
                    for hit in [self.tab.fx.get(id(row))] if hit and hit[0] is row for ci, e in hit[1].items()]
        cur = self.sheet.get_currently_selected()
        at = (cur.row, self.cols[cur.column], cur.type_) if cur and cur.column < n else None
        self._show(cols, data, orig)
        self.tab.fx = {}
        for gi, col, e in formulas:  # same rows, new row lists: move the formulas over
            if col in self.cols and gi < len(self.sheet.MT.data):
                row = self.sheet.MT.data[gi]
                self.tab.fx.setdefault(id(row), (row, {}))[1][self.cols.index(col)] = e
        if at and at[1] in self.cols:  # same cell / column stays selected
            c = self.cols.index(at[1])
            if at[2] == "columns":
                self.sheet.select_column(c)
            else:
                self.sheet.select_cell(at[0], c)

    def _label_display(self):
        """Shown text of lookup cells: "6 · Valoren" in same mode, a ⚠ before a
        value missing from the lookup. Render only - the cell keeps the id."""
        mark = self.settings.get("mismatch", "both") in ("mark", "both")
        self._sub_formats()
        disp = {}
        for i, c in enumerate(self.cols):
            if c.of is not None or (c.choices is None and not c.big):  # label columns hold shown text
                continue
            mode = self._label_mode(c)
            same = mode in ("same", "value") and c.labelled
            if not same and not mark:
                continue
            fmt = self._label_fmt(mode)

            def fn(v, c=c, same=same, fmt=fmt):
                known = c.in_lookup(v) if v else None
                if known is None:
                    return v
                if known:
                    label = label_of(c, v, mark) if same else None
                    return fmt.format(id=v, label=label) if label else v
                return WARN + v if mark else v

            disp[i] = fn
        self.sheet.disp = disp

    def _sync_labels(self, rows):
        """Label columns follow their id column (typed, pasted, filled)."""
        mark = self.settings.get("mismatch", "both") in ("mark", "both")
        data = self.sheet.MT.data
        for vi, c in enumerate(self.cols):
            if c.of is None or c.of not in self.cols:
                continue
            si = self.cols.index(c.of)
            for gi, r in enumerate(rows):
                want = label_text(c.of, r[si], mark, c.part)
                if r[vi] != want:
                    r[vi] = data[gi][vi] = want

    def label_display_menu(self):
        """Header menu: how this lookup column shows labels (overrides Settings)."""
        i = self._col()
        if i is None:
            return
        col = self.cols[i].of or self.cols[i]
        if not col.labelled:
            return self.msg(f"{col.name} has no lookup labels to show", error=True)
        pick = LabelModeDialog(self, col).result
        if pick is None:
            return
        modes = self.settings.setdefault("label_modes", {})
        key = col_key(self.database, self.table, col.name)
        if pick == "default":
            modes.pop(key, None)
        else:
            modes[key] = pick
        save_settings(self.settings)
        self._relabel()

    def lookup_edit(self):
        """Header menu: choose the lookup of a column - its table, key and the
        column shown beside the key - or turn it off, or back to automatic.
        Remembered per column."""
        i = self._col()
        if i is None:
            return
        col = self.cols[i]
        if col.of is not None and col.part is not None:  # a label field: its own lookup
            top = col.of
            col = top.subs.get(col.part) or Col(col.name, "", False, False, False, external=True)
        elif col.external or col.readonly or not self.table:
            return self.msg(f"{col.name} is read-only: no lookup", error=True)
        else:
            top = col
        if not self.table:
            return self.msg("Lookups belong to a table: load one first", error=True)
        dlg = LookupDialog(self, col)
        if dlg.result is None and not dlg.subs_changed:
            return
        if dlg.result is not None:
            self.save_lookup(col.name, dlg.result, dlg.show)
        col = top
        spec = self._lookup_spec()
        if spec is None:
            return
        tab, database, table, db, prefs = self.tab, self.database, self.table, self.db, self._lookup_prefs()
        cols = [c for c in self.cols if c.name == col.name and not c.external]

        def work():
            db.lookups(database, table, cols, spec, prefs)

        def done(_, secs):
            if self.tab is tab:
                self._relabel()
                self.msg(f"{col.name}: " + (f"lookup {col.lookup}, labels read per value" if col.big else
                                            "no dropdown" if col.choices is None else
                                            f"{len(col.choices):,} values from {col.lookup}"))

        self.bg("Reading lookup", work, done)

    def save_lookup(self, name: str, pick, show: str):
        """Remember a lookup choice ("auto", {"off"}, a table) and its display for column / field `name`."""
        key = col_key(self.database, self.table, name)
        modes = self.settings.setdefault("label_modes", {})
        if show == "default":
            modes.pop(key, None)
        else:
            modes[key] = show
        lookups = self.settings.setdefault("lookups", {})
        if pick == "auto":
            lookups.pop(key, None)
        else:
            lookups[key] = pick
        save_settings(self.settings)

    def scratch_add(self):
        """A column of your own right of the selected one: type notes, work things out
        with formulas. Kept on this sheet; Insert / Update never send it."""
        i = self._col()
        if i is None:
            return self.msg("Load a table or run a query first", error=True)
        taken = {c.label.lower() for c in self.cols}
        k = 1
        while f"scratch{k}" in taken:
            k += 1
        name = simpledialog.askstring("Add scratch column", "Name of the new column\n"
                                      "(on this sheet only, never written to the database):",
                                      initialvalue=f"Scratch{k}", parent=self.root)
        if not name or not name.strip():
            return
        name = name.strip()
        if name.lower() in taken:
            return self.msg(f"There is already a column {name}", error=True)
        col = Col(name, "", False, False, False, external=True)
        col.scratch, col.readonly = True, False
        self._relabel(add=(i + 1, col))
        c = self.cols.index(col)
        self.sheet.select_column(c)
        self.sheet.see(0, c)
        self.msg(f"{name}: scratch column, kept on this sheet only and never written")

    def _scratch_col(self) -> "Col | None":
        i = self._col()
        if i is None or not self.cols[i].scratch:
            self.msg("Select a scratch column first", error=True)
            return None
        return self.cols[i]

    def scratch_rename(self):
        col = self._scratch_col()
        if col is None:
            return
        name = simpledialog.askstring("Rename scratch column", "New name:", initialvalue=col.label,
                                      parent=self.root)
        if not name or not name.strip() or name.strip() == col.label:
            return
        if name.strip().lower() in {c.label.lower() for c in self.cols if c is not col}:
            return self.msg(f"There is already a column {name.strip()}", error=True)
        col.name = col.label = name.strip()
        self._relabel()

    def scratch_remove(self):
        col = self._scratch_col()
        if col is None:
            return
        filled = sum(1 for r in self._rows() if r[self.cols.index(col)].strip())
        if filled and not messagebox.askyesno("Remove scratch column",
                                              f"{col.label} has {filled:,} filled cells. Remove it?",
                                              parent=self.root):
            return
        self._relabel(drop=col)
        self.msg(f"{col.label} removed")

    def find_column(self):
        """Ctrl+Shift+F: jump to a column by (part of) its name."""
        if self.cols:
            ColumnFind(self)

    def lookup_view(self):
        """Header menu: every allowed value of the column with its label."""
        i = self._col()
        if i is None:
            return
        col = self.cols[i]
        if col.choices is None and not col.big:
            return self.msg(f"{col.name} has no lookup list" + (f" ({col.lookup})" if col.lookup else ""),
                            error=True)
        LookupView(self, col)

    def put_value(self, col: "Col", value: str):
        """Write `value` into the selected cells of `col` (from LookupView)."""
        c = self.cols.index(col)
        rows = sorted({self._dr(r) for r, cc in self.sheet.get_selected_cells() if cc == c})
        if not rows:
            cur = self.sheet.get_currently_selected()
            rows = [self._dr(cur.row)] if cur and cur.column == c else []
        if not rows:
            return self.msg(f"Select cells in {col.name} first", error=True)
        for r in rows:
            self.sheet.set_cell_data(r, c, value)
        self.sheet.redraw()
        self._on_select()
        self._schedule_marks()

    def _dropdowns(self):
        """Lookup columns get a dropdown of their allowed values (blank = NULL first)."""
        for i in list(self.sheet.MT.col_options):
            if "dropdown" in self.sheet.MT.col_options[i]:
                self.sheet.del_dropdown(self.sheet.span(num2alpha(i), header=False), redraw=False)
        for i, c in enumerate(self.cols):
            if c.choices is not None and not (c.readonly or c.audit):
                self.sheet.dropdown(self.sheet.span(num2alpha(i), header=False), values=["", *c.choice_list],
                                    edit_data=False, validate_input=False, redraw=False,
                                    search_function=choice_search)

    def _fit_columns(self, data):
        ch = tkfont.Font(family=MONO[0], size=10).measure("0")
        hd = tkfont.Font(family=FONT[0], size=10)
        widths = []
        for i, c in enumerate(self.cols):
            longest = max((len(r[i]) for r in data[:300]), default=0)
            if c.choices is not None:  # room for the dropdown entries and the arrow
                longest = max([longest, *(len(e) + 3 for e in c.choice_list[:300])])
            widths.append(min(320, max(80, hd.measure(c.label) + 28, longest * ch + 16)))
        self.sheet.set_column_widths(widths)

    # grid helpers ----------------------------------------------------------
    def _rows(self) -> list[list[str]]:
        """Grid data as text, n columns + the hidden row id. Reads tksheet's own
        lists: its per-cell getter is ~1 µs a cell, seconds on a big paste."""
        n = len(self.cols) + 1
        out = []
        for r in self.sheet.MT.data:
            if len(r) == n and None not in r:  # the app only ever stores text
                out.append(r[:])
                continue
            r = [("" if v is None else str(v)) for v in r[:n]]
            r += [""] * (n - len(r))
            out.append(r)
        return out

    def _twins(self) -> list[list[int]]:
        """Grid columns that are the same table column, e.g. IssueDate in
        `SELECT IssueDate, * FROM t`. First index of each group is the one written."""
        groups: dict[str, list[int]] = {}
        for i, c in enumerate(self.cols):
            if not c.external:
                groups.setdefault(c.name.lower(), []).append(i)
        return [g for g in groups.values() if len(g) > 1]

    def _sync_twins(self, rows) -> bool:
        """Copy an edit made in one copy of a column to its other copies.
        The copy that differs from the loaded value wins; for new rows, the
        filled-in copy wins. Returns True when the grid was changed."""
        n, changed = len(self.cols), False
        for group in self._twins():
            for gi, r in enumerate(rows):
                vals = [r[i] for i in group]
                if len(set(vals)) == 1:
                    continue
                o = self.orig.get(r[n])
                if o is not None:
                    pick = next((v for i, v in zip(group, vals) if v != o[i]), vals[0])
                else:
                    pick = next((v for v in vals if v != ""), "")
                for i in group:
                    if r[i] != pick:
                        r[i] = pick
                        self.sheet.set_cell_data(gi, i, pick, redraw=False)
                        changed = True
        return changed

    def _diff(self, rows):
        """-> (new [(grid_row, values)], edited [(grid_row, rid, changed_idx)])"""
        n = len(self.cols)
        mirrors = {i for g in self._twins() for i in g[1:]}  # written through the first copy
        new, edited = [], []
        for gi, r in enumerate(rows):
            rid, vals = r[n], r[:n]
            if rid and rid in self.orig:
                o = self.orig[rid]
                ch = [i for i in range(n)
                      if nz(vals[i]) != nz(o[i]) and not self.cols[i].readonly and i not in mirrors
                      and not self.cols[i].scratch]
                if ch:
                    edited.append((gi, rid, ch))
            elif any(nz(v.strip()) is not None for v, c in zip(vals, self.cols) if not c.scratch):
                new.append((gi, vals))
        return new, edited

    def _schedule_marks(self):
        if self._mark_after:
            self.root.after_cancel(self._mark_after)
        self._mark_after = self.root.after(200, self._marks)

    def _marks(self):
        self._mark_after = None
        t = self.tab
        if t.label is not None and t.label.winfo_exists():
            t.label.configure(text=self._tab_text(t))
        if not self.cols:
            self.bad = {}
            self.l_status.configure(text="")
            for b, word in ((self.b_insert, "Insert"), (self.b_update, "Update")):
                b.configure(text=word)
                b.state(["disabled"])
            return
        rows = self._rows()
        self._sync_twins(rows)
        self._sub_formats()
        self._sync_labels(rows)
        self._label_display()
        new, edited = self._diff(rows)
        self._formulas(rows, new, edited)
        self._big_lookups(rows, new, edited)
        self.sheet.dehighlight_all(redraw=False)
        ro = [i for i, c in enumerate(self.cols) if c.readonly or c.audit]
        if ro:
            self.sheet.highlight_columns(ro, bg=C["readonly"], redraw=False)
        for gi, _ in new:
            self.sheet.highlight_rows([gi], bg=C["new"], redraw=False)
        twin_of = {g[0]: g[1:] for g in self._twins()}
        cells = [(gi, x) for gi, _, ch in edited for ci in ch for x in [ci, *twin_of.get(ci, [])]]
        for i, c in enumerate(self.cols):  # keys and lookups stand out in the header
            if c.scratch:
                self.sheet.highlight_cells(row=0, column=i, canvas="header", bg=C["new"], redraw=False)
            elif c.is_pk and not c.external:
                self.sheet.highlight_cells(row=0, column=i, canvas="header", fg=C["accent"], redraw=False)
            elif c.choices is not None or c.big or c.fk:
                self.sheet.highlight_cells(row=0, column=i, canvas="header", fg=C["link"], redraw=False)
        filtered = [i for i, c in enumerate(self.cols) if c.label in self.filters]
        for i in filtered:
            self.sheet.highlight_cells(row=0, column=i, canvas="header", bg=C["accent_soft"],
                                       fg=C["accent"], redraw=False)
        if cells:
            self.sheet.highlight_cells(cells=cells, bg=C["edited"], redraw=False)
        if self.settings.get("mismatch", "both") in ("red", "both"):  # values missing from a lookup
            miss = [(gi, ci) for ci, c in enumerate(self.cols) if c.of is None and (c.choices is not None or c.big)
                    for gi, r in enumerate(rows) if r[ci] and c.in_lookup(r[ci]) is False]
            if miss:
                self.sheet.highlight_cells(cells=miss, fg=C["error"], redraw=False)
        bad = self.bad = self._bad_cells(rows, new, edited)
        if bad:
            self.sheet.highlight_cells(cells=list(bad), bg=C["old"], fg=C["error"], redraw=False)
        self.sheet.redraw()
        writable = bool(self.table) and any(not c.readonly and not c.scratch for c in self.cols)
        bad_rows = {r for r, _ in bad}
        for b, items, word in ((self.b_insert, new, "Insert"), (self.b_update, edited, "Update")):
            n = len(items)
            ok = n and writable and not self._busy and not any(it[0] in bad_rows for it in items)
            b.configure(text=f"{word} {n:,}" if n else word)
            b.state(["!disabled"] if ok else ["disabled"])
        if not writable:
            parts = [self.tab.server, self.database, "query result", f"{len(self.orig):,} rows",
                     "read-only: no table of its own to edit"]
        else:
            parts = [self.tab.server, f"{self.database}.{self.table}",
                     "from query" if self.from_query else "",
                     f"filtered {self._shown_rows():,} of {len(rows):,}" if self.filters else "",
                     f"{len(self.orig):,} loaded", f"{len(edited):,} edited", f"{len(new):,} new",
                     f"{len(bad):,} invalid" if bad else "",
                     "key " + (", ".join(k.name for k in self.keys) or "none")]
        self.l_status.configure(text="  ·  ".join(p for p in parts if p))
        self._top_up(rows)

    # live formulas ----------------------------------------------------------
    # A cell holds its result (what Insert / Update write); its formula lives in
    # tab.fx, keyed by the row's list object so sort / filter / new rows keep it.
    def _fx_get(self, gi: int, ci: int) -> dict | None:
        data = self.sheet.MT.data
        if gi >= len(data):
            return None
        hit = self.tab.fx.get(id(data[gi]))
        return hit[1].get(ci) if hit and hit[0] is data[gi] else None

    def fx_text(self, gi: int, ci: int) -> str:
        """What a cell is: its formula, else its value (fx bar, editor, fill, copy)."""
        e = self._fx_get(gi, ci)
        if e:
            return e["f"]
        v = self.sheet.get_cell_data(gi, ci) if gi < self.sheet.get_total_rows() else ""
        return "" if v is None else str(v)

    def _retry_formula(self, gi: int, ci: int):
        """Entered again: evaluate again (the error may have been fixed on the server)."""
        e = self._fx_get(gi, ci)
        if e:
            e["err"] = e["sig"] = None
        self.tab.fx_hold = False

    def _formulas(self, rows, new, edited):
        """Keep formulas in step with the grid and evaluate those that need it.
        Typed, pasted and filled "=…" cells become formulas (loaded data never does);
        a value typed over a formula ends it. A formula is evaluated when it is new or
        a cell it reads changed; one that reads a formula waits for that one first.
        Cells sharing a shape (a filled-down formula) go to the server as one query."""
        t, data, n = self.tab, self.sheet.MT.data, len(self.cols)
        pos = {id(r): gi for gi, r in enumerate(data)}
        for k in [k for k, (row, _) in t.fx.items() if pos.get(k) is None or data[pos[k]] is not row]:
            del t.fx[k]  # row deleted, or the grid reloaded
        ok = {ci for ci, c in enumerate(self.cols) if not (c.external or c.readonly or c.audit)}
        entered = [(gi, ci) for gi, _ in new for ci in ok] + [(gi, ci) for gi, _, ch in edited for ci in ch if ci in ok]
        local = [ci for ci, c in enumerate(self.cols) if c.scratch]
        entered += [(gi, ci) for ci in local for gi, r in enumerate(rows) if is_formula(r[ci])]
        for gi, ci in entered:
            v = rows[gi][ci]
            if is_formula(v):
                cells = t.fx.setdefault(id(data[gi]), (data[gi], {}))[1]
                e = cells.get(ci)
                if e is None or (v != e["f"] and v != e["res"]):
                    cells[ci] = {"f": v, "res": None, "sig": None, "err": None}
        fcells: dict[tuple[int, int], dict] = {}
        for k, (row, cells) in list(t.fx.items()):
            gi = pos[k]
            for ci, e in list(cells.items()):
                cur = rows[gi][ci] if ci < n else ""
                if cur == e["f"] or (e["res"] is not None and cur == e["res"]):
                    fcells[gi, ci] = e
                else:  # a value typed or pasted over it
                    del cells[ci]
            if not cells:
                del t.fx[k]
        t.fx_why = {}
        if not fcells:
            return
        # what each formula needs, in dependency order (iterative: fill-down chains are long)
        status: dict[tuple[int, int], str] = {}  # ok / eval / wait / err
        plan: dict[str, list] = {}
        for start in fcells:
            if start in status:
                continue
            stack, onpath = [(start, None)], set()
            while stack:
                cell, info = stack[-1]
                if info is None:
                    onpath.add(cell)
                    e = fcells[cell]
                    try:
                        info = fx_parse(e["f"], cell[0], self.cols, len(rows))
                    except FxError as x:
                        info = x
                    deps = [] if isinstance(info, FxError) else \
                        [d for r in info[1] for d in ([(r[2], r[1])] if r[0] == "cell" else r[2]) if d in fcells]
                    stack[-1] = (cell, (info, deps))
                    nxt = next((d for d in deps if d not in status and d not in onpath), None)
                    if nxt is not None:
                        stack.append((nxt, None))
                        continue
                else:
                    info, deps = info
                    nxt = next((d for d in deps if d not in status and d not in onpath), None)
                    if nxt is not None:
                        stack.append((nxt, None))
                        continue
                stack.pop()
                onpath.discard(cell)
                status[cell] = self._fx_status(cell, fcells[cell], info, deps, status, rows, plan)
        if plan and not self._busy and not t.fx_hold and t.db.cn is not None:
            self._fx_run(plan)

    def _fx_status(self, cell, e, info, deps, status, rows, plan) -> str:
        t = self.tab
        gi, ci = cell
        if isinstance(info, FxError):
            t.fx_why[cell] = f"Formula: {info}"
            return "err"
        if any(d == cell or status.get(d) is None for d in deps):  # still on the path: a loop
            t.fx_why[cell] = "Formula: circular reference"
            return "err"
        bad = next((d for d in deps if status[d] == "err"), None)
        if bad is not None:
            t.fx_why[cell] = f"Formula: {num2alpha(bad[1])}{bad[0] + 1} has an error"
            return "err"
        if any(status[d] in ("eval", "wait") for d in deps):
            t.fx_why[cell] = "Formula: waiting for the cells it reads"
            return "wait"
        pieces, refs, shape = info

        def val(r, c):
            v = rows[r][c]
            return None if v == "" or v.strip().upper() == "NULL" else v

        values = [val(r[2], r[1]) if r[0] == "cell" else [val(rr, cc) for rr, cc in r[2]] for r in refs]
        sig = (e["f"], json.dumps(values))
        cur = rows[gi][ci]
        if e["err"] and e["sig"] == sig:
            t.fx_why[cell] = f"Formula failed: {e['err']}"
            return "err"
        if cur == e["f"] or e["sig"] != sig:
            col = self.cols[ci]
            key = f"{shape}\t{sql_type(col)}"
            if key not in plan:
                plan[key] = [fx_sql(pieces, refs, self.cols, sql_type(col)), []]
            item = {"i": len(plan[key][1])}
            item.update({f"p{k}": v for k, v in enumerate(values)})
            plan[key][1].append((item, self.sheet.MT.data[gi], ci, sig, cur))
            t.fx_why[cell] = ("Formula: connect to evaluate it" if t.db.cn is None else
                              "Formula: not evaluated, enter it again" if t.fx_hold and not self._busy
                              else "Formula: calculating …")
            return "eval"
        return "ok"

    def _fx_run(self, plan: dict):
        t, db, database = self.tab, self.db, self.database
        jobs = []
        for key, (sql, cells) in plan.items():
            err = check_select(sql)
            jobs.append((sql, cells, err))
        t.fx_hold = True  # stays set if cancelled, so Cancel is not undone by the next redraw

        def work():
            return [(cells, RuntimeError(err)) if err else (cells, db.eval_fx(database, sql, [c[0] for c in cells]))
                    for sql, cells, err in jobs]

        def done(res, secs):
            t.fx_hold = False
            data = self.sheet.MT.data
            pos = {id(r): gi for gi, r in enumerate(data)}
            n_ok, failed = 0, []
            for cells, out in res:
                for item, row, ci, sig, cur in cells:
                    gi = pos.get(id(row))
                    e = self._fx_get(gi, ci) if gi is not None else None
                    if e is None or data[gi][ci] != cur or e["f"] != sig[0]:
                        continue  # changed while the server worked: the next round redoes it
                    v = out if isinstance(out, Exception) else out.get(item["i"])
                    e["sig"] = sig
                    if isinstance(v, Exception):
                        e["err"] = clean_err(v)
                        failed.append(f"{num2alpha(ci)}{gi + 1} {e['f']}: {e['err']}")
                        continue
                    v = "" if v is None else v
                    e["err"], e["res"] = None, v
                    if data[gi][ci] != v:
                        self.sheet.set_cell_data(gi, ci, v, redraw=False)
                    n_ok += 1
            if failed:
                more = f"\n… and {len(failed) - 5:,} more" if len(failed) > 5 else ""
                self.msg("Formula failed: " + "\n".join(failed[:5]) + more, error=True)
            else:
                self.msg(f"{n_ok:,} formula cell{'s' * (n_ok != 1)} calculated ({secs:.1f}s)")
            self.sheet.redraw()
            self._on_select()

        self.bg("Calculating formulas", work, done, cancel=True)

    def _big_lookups(self, rows, new, edited):
        """Big lookups (too many rows to list): read the label fields of the values in
        the grid, in the background. An entered value that is no key is looked up by
        its label and replaced by the key of the one row it names."""
        t = self.tab
        if self._busy or t.lk_hold or t.db.cn is None:
            return
        entered = {(gi, ci) for gi, _ in new for ci in range(len(self.cols))}
        entered |= {(gi, ci) for gi, _, ch in edited for ci in ch}
        todo = []
        for ci, c in enumerate(self.cols):
            if not c.big or c.of is not None:
                continue
            vals = {r[ci] for r in rows} - {""}
            ask = [v for v in vals if v not in c.key_fields and v not in c.missing and not is_formula(v)
                   and v.strip().upper() != "NULL"]
            typed = {rows[gi][cc] for gi, cc in entered if cc == ci}
            if ask:
                todo.append((c, ask, typed))
            stack = [(c, vals)]  # sub lookups: the field values of the keys in use
            while stack:
                owner, used = stack.pop()
                for part, sub in owner.subs.items():
                    fvs = {str(owner.key_fields[v][part]) for v in used
                           if v in owner.key_fields and part < len(owner.key_fields[v])} - {"", "None"}
                    if sub.big:
                        sub_ask = [v for v in fvs if v not in sub.key_fields and v not in sub.missing]
                        if sub_ask:
                            todo.append((sub, sub_ask, set()))
                    stack.append((sub, fvs))
        if not todo:
            return
        db = self.db
        t.lk_hold = True  # stays set if cancelled or failed

        def work():
            out = []
            for c, ask, typed in todo:
                found, missing = db.big_labels(c, ask)
                named = {v: db.big_resolve(c, v) for v in missing if v in typed}
                out.append((c, found, missing, named))
            return out

        def done(out, secs):
            t.lk_hold = False
            now = self._rows()
            n = 0
            for c, found, missing, named in out:
                for v, vals in found.items():
                    c.key_fields[v], c.key_label[v] = vals, LABEL_SEP.join(str(x) for x in vals if x)
                c.missing |= missing
                if c not in self.cols:
                    continue
                ci = self.cols.index(c)
                for text, res in named.items():
                    if isinstance(res, str):
                        c.nomatch[text] = res
                        continue
                    key, vals = res
                    c.key_fields[key], c.key_label[key] = vals, LABEL_SEP.join(str(x) for x in vals if x)
                    for gi, r in enumerate(now):
                        if r[ci] == text and (gi, ci) in entered:
                            self.sheet.set_cell_data(gi, ci, key, redraw=False)
                            n += 1
            found = sum(len(f) for _, f, _, _ in out)
            self.msg(f"{n:,} name{'s' * (n != 1)} replaced by the lookup key" if n else
                     f"Lookup labels read: {found:,} found · {secs:.1f}s")
            self.sheet.redraw()

        self.bg("Reading lookup labels", work, done, cancel=True)

    def _bad_cells(self, rows, new, edited) -> dict[tuple[int, int], str]:
        """Cells that would fail on the server, with the reason: every filled cell of
        a new row, and the changed cells of an edited row."""
        bad = {}
        checked = [(ci, col) for ci, col in enumerate(self.cols)
                   if not (col.external or col.readonly or col.audit)]
        for gi, vals in new:
            for ci, col in checked:
                v = vals[ci]
                if v and (err := cell_error(col, v)):  # blank is always fine
                    bad[gi, ci] = err
        for gi, _, ch in edited:
            for ci in ch:
                err = cell_error(self.cols[ci], rows[gi][ci])
                if err:
                    bad[gi, ci] = err
        for gi, ci in [(gi, ci) for gi, _ in new for ci in range(len(self.cols))] + \
                [(gi, ci) for gi, _, ch in edited for ci in ch]:
            col, v = self.cols[ci], rows[gi][ci]
            if col.big and col.of is None and v and v.strip().upper() != "NULL" and not is_formula(v) \
                    and col.in_lookup(v) is None and v not in col.nomatch:
                bad[gi, ci] = "Checking the lookup …"
        bad.update(self.tab.fx_why)  # formulas not calculated yet, or failed
        return bad

    def _spare(self, row: list[str]) -> bool:
        """An empty row for typing (no DB origin, nothing entered), not data."""
        n = len(self.cols)
        return not row[n] and not any(v.strip() for v in row[:n])

    def _top_up(self, rows):
        """Keep empty rows under the data, as in a spreadsheet, for typing new records."""
        n = len(self.cols)
        spare = 0
        for r in reversed(rows):
            if any(v.strip() for v in r[:n]):
                break
            spare += 1
        if spare < SPARE_ROWS // 2:
            self.sheet.insert_rows(SPARE_ROWS - spare, redraw=True)

    def _ensure_keys(self) -> bool:
        if self.keys:
            return True
        table_pk = [c.name for c in self.table_cols if c.is_pk]
        if table_pk:  # key exists on the table but the query left it out
            self.msg(f"Select {', '.join(table_pk)} from {self.table} once in the query to update rows",
                     error=True)
            return False
        picked = KeyDialog(self.root, [c.name for c in self.cols if not c.external]).result
        if not picked:
            self.msg("Update needs key columns", error=True)
            return False
        self.keys = [c for c in self.cols if c.name in picked]
        self._marks()
        return True

    # paste / import / export -----------------------------------------------
    def _append_mapped(self, header: list[str], body: list[list[str]], source: str):
        self._fx_flush()
        by_name = {c.name.lower(): i for i, c in enumerate(self.cols)}
        by_name.update({c.label.lower(): i for i, c in enumerate(self.cols)})
        idx = [by_name.get(h.strip().lower()) for h in header]
        unknown = [h for h, i in zip(header, idx) if i is None and h.strip()]
        n = len(self.cols)
        add = []
        for r in body:
            if not any(str(v).strip() for v in r):
                continue
            row = [""] * (n + 1)
            for v, i in zip(r, idx):
                if i is not None:
                    row[i] = fmt(v) if not isinstance(v, str) else v
            add.append(row)
        if not any(i is not None for i in idx):
            return self.msg(f"No {source} headers match columns of {self.table}", error=True)
        data = [r for r in self._rows() if any(v.strip() for v in r)] + add
        self.filters = {}  # show the rows just added
        self.sheet.display_rows("all", all_rows_displayed=True)
        self.sheet.set_sheet_data(data, reset_col_positions=False)
        self.sheet.display_columns(list(range(n)), all_columns_displayed=False,
                                  reset_col_positions=False)
        self.sheet.readonly_columns([i for i, c in enumerate(self.cols) if c.readonly or c.audit], True)
        self._marks()
        note = f" · ignored columns: {', '.join(unknown)}" if unknown else ""
        self.msg(f"Added {len(add):,} rows from {source}{note}")

    def paste_with_headers(self):
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            return self.msg("Clipboard is empty", error=True)
        delim = "\t" if "\t" in text.split("\n", 1)[0] else ","
        rows = list(csv.reader(text.splitlines(), delimiter=delim))
        if len(rows) < 2:
            return self.msg("Clipboard needs a header row and at least one data row", error=True)
        self._append_mapped(rows[0], rows[1:], "clipboard")

    def import_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV / Excel", "*.csv *.tsv *.txt *.xlsx"),
                                                     ("All files", "*.*")])
        if not path:
            return
        try:
            if path.lower().endswith(".xlsx"):
                import openpyxl  # optional dependency
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                rows = [["" if v is None else v for v in r] for r in wb.active.iter_rows(values_only=True)]
            else:
                with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
                    sample = f.read(4096); f.seek(0)
                    delim = "\t" if "\t" in sample.split("\n", 1)[0] else ","
                    rows = list(csv.reader(f, delimiter=delim))
        except ImportError:
            return self.msg("Install openpyxl for .xlsx import: pip install openpyxl", error=True)
        except Exception as e:  # noqa: BLE001
            return self.msg(f"Cannot read file: {e}", error=True)
        if len(rows) < 2:
            return self.msg("File needs a header row and data rows", error=True)
        self._append_mapped([str(h) for h in rows[0]], rows[1:], Path(path).name)

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            initialfile=f"{self.table.replace('.', '_')}.csv")
        if not path:
            return
        self._fx_flush()
        n = len(self.cols)
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow([c.name for c in self.cols])
            w.writerows(r[:n] for r in self._rows() if any(v.strip() for v in r[:n]))
        self.msg(f"Exported to {path}")

    def _dry_csv(self, action: str, header: list[str], lines: list[list]) -> Path:
        DRY_RUN_DIR.mkdir(exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        p = DRY_RUN_DIR / f"{self.database}_{self.table.replace('.', '_')}_{action}_{stamp}.csv"
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerows(lines)
        return p

    # writes ----------------------------------------------------------------
    def insert(self):
        self._fx_flush()
        new, _ = self._diff(self._rows())
        if not new:
            return self.msg("No new rows. New rows are the green ones without a DB origin.")
        dry = self.v_dry.get()
        sql, data_cols = build_insert(self.table, self.cols, self.table_cols)
        pos = [self.cols.index(c) for c in data_cols]
        payloads = [(gi, {f"v{j}": nz(vals[p]) for j, p in enumerate(pos)}) for gi, vals in new]
        database = self.database
        if self.settings["review"]:
            verb = "Dry run" if dry else "Insert"
            dlg = ReviewDialog(
                self.root, "Review insert",
                f"{len(new):,} new rows into {database}.{self.table}. Every value below is new; "
                "blank = NULL. Audit columns are filled on the server.",
                ["Grid row", *[c.name for c in data_cols]],
                [[gi + 1, *[vals[i] for i in pos]] for gi, vals in new],
                {i + 1: "new" for i in range(len(data_cols))},
                f"{verb} {len(new):,} row{'s' if len(new) != 1 else ''}")
            if not dlg.ok:
                return self.msg("Insert cancelled")

        db = self.db

        def work():
            return db.write(database, [(sql, payloads, False)], self.user, dry)

        def done(_n, secs):
            if dry:
                p = self._dry_csv("insert", ["grid_row", *[c.name for c in data_cols]],
                                  [[gi + 1, *[vals[i] for i in pos]] for gi, vals in new])
                return self.msg(f"Dry run · {len(new):,} rows would insert, validated and rolled back "
                                f"· {secs:.1f}s · {p}")
            drop = {gi for gi, _ in new}
            keep = [r for gi, r in enumerate(self._rows()) if gi not in drop]
            self.sheet.display_rows("all", all_rows_displayed=True)
            self.sheet.set_sheet_data(keep, reset_col_positions=False)
            self.sheet.display_columns(list(range(len(self.cols))), all_columns_displayed=False,
                                       reset_col_positions=False)
            self._marks()
            self.msg(f"Inserted {len(new):,} rows into {database}.{self.table} · {secs:.1f}s · "
                     "Load to see them")

        self.bg(f"Inserting {len(new):,} rows", work, done)

    def update(self):
        self._fx_flush()
        rows = self._rows()
        _, edited = self._diff(rows)
        if not edited:
            return self.msg("No edited cells")
        if not self._ensure_keys():
            return
        dry = self.v_dry.get()
        kpos = [self.cols.index(k) for k in self.keys]
        all_edited = edited
        edited, conflicts = dedupe_edits(edited, rows, self.orig, kpos,
                                           [c.label for c in self.cols])
        if conflicts:
            return self.msg("Same row edited differently in several grid rows: " + "; ".join(conflicts),
                            error=True)
        groups: dict[tuple, list] = {}
        for gi, rid, ch in edited:
            groups.setdefault(tuple(ch), []).append((gi, rid))
        plan = []
        for ch, members in groups.items():
            changed = [self.cols[i] for i in ch]
            sql = build_update(self.table, self.table_cols, self.keys, changed)
            payloads = []
            for gi, rid in members:
                o = self.orig[rid]
                p = {f"k{j}": o[kp] for j, kp in enumerate(kpos)}
                p.update({f"v{j}": nz(rows[gi][ci]) for j, ci in enumerate(ch)})
                payloads.append((gi, p))
            plan.append((sql, payloads, True))
        database = self.database
        lines = [[gi + 1, *[self.orig[rid][kp] for kp in kpos], self.cols[ci].name,
                  self.orig[rid][ci], rows[gi][ci]] for gi, rid, ch in edited for ci in ch]
        if self.settings["review"]:
            nk = len(self.keys)
            verb = "Dry run" if dry else "Update"
            dlg = ReviewDialog(
                self.root, "Review update",
                f"{len(edited):,} rows, {len(lines):,} cells in {database}.{self.table}. "
                "Only these cells are written; blank = NULL.",
                ["Grid row", *[k.name for k in self.keys], "Column", "Old value", "New value"],
                lines, {nk + 2: "old", nk + 3: "new"},
                f"{verb} {len(edited):,} row{'s' if len(edited) != 1 else ''}")
            if not dlg.ok:
                return self.msg("Update cancelled")

        db = self.db

        def work():
            return db.write(database, plan, self.user, dry)

        def done(n, secs):
            if dry:
                p = self._dry_csv("update", ["grid_row", *[k.name for k in self.keys], "column",
                                             "old", "new"], lines)
                return self.msg(f"Dry run · {len(edited):,} rows would update, validated and rolled "
                                f"back · {secs:.1f}s · {p}")
            for gi, rid, _ in all_edited:  # duplicates of a deduped row are saved too
                self.orig[rid] = rows[gi][:len(self.cols)]
            self._marks()
            self.msg(f"Updated {n:,} rows in {database}.{self.table} · {secs:.1f}s")

        self.bg(f"Updating {len(edited):,} rows", work, done)


for _name in TAB_STATE:  # App.<name> is the current tab's <name>
    setattr(App, _name, property(lambda self, n=_name: getattr(self.tab, n),
                                 lambda self, v, n=_name: setattr(self.tab, n, v)))


class KeyDialog(tk.Toplevel):
    """Pick key columns when the table has no primary key or unique index."""

    def __init__(self, parent, names: list[str]):
        super().__init__(parent)
        self.title("Key columns")
        self.configure(bg=C["surface"])
        self.result: list[str] = []
        self.transient(parent)
        ttk.Label(self, text="No primary key found. Pick the columns that identify a row.",
                  padding=(16, 16, 16, 8)).pack(anchor="w")
        self.lb = tk.Listbox(self, selectmode="multiple", height=14, font=MONO, bg=C["surface"],
                             fg=C["text"], selectbackground=C["accent_soft"], selectforeground=C["text"],
                             highlightthickness=1, highlightbackground=C["line"], relief="flat")
        for n in names:
            self.lb.insert("end", n)
        self.lb.pack(fill="both", expand=True, padx=16)
        bar = ttk.Frame(self, padding=16)
        bar.pack(fill="x")
        ttk.Button(bar, text="Use as key", style="Primary.TButton", command=self.ok).pack(side="right")
        ttk.Button(bar, text="Cancel", command=self.destroy).pack(side="right", padx=8)
        modal(self)
        self.wait_window()

    def ok(self):
        self.result = [self.lb.get(i) for i in self.lb.curselection()]
        self.destroy()


def select_all(e):
    w = e.widget
    if isinstance(w, tk.Text):
        w.tag_add("sel", "1.0", "end-1c")
        w.mark_set("insert", "end-1c")
    else:
        w.selection_range(0, "end")
        w.icursor("end")
    return "break"


def paste_replace(e):
    """Paste over the selected text, as on Windows (Tk on Linux inserts beside it)."""
    w = e.widget
    try:
        text = w.clipboard_get()
    except tk.TclError:
        return "break"
    if isinstance(w, ttk.Widget):
        if w.instate(["readonly"]) or w.instate(["disabled"]):
            return "break"
    elif str(w.cget("state")) in ("readonly", "disabled"):
        return "break"
    try:
        w.delete("sel.first", "sel.last")
    except tk.TclError:  # nothing selected
        pass
    w.insert("insert", text)
    if isinstance(w, tk.Text):
        w.see("insert")
    return "break"


class MessageLog(tk.Toplevel):
    """Full text of recent errors and long messages, newest first, copyable."""

    def __init__(self, app: "App"):
        super().__init__(app.root)
        self.title("Messages")
        self.configure(bg=C["surface"])
        self.transient(app.root)
        self.geometry("760x420")
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both", expand=True)
        t = tk.Text(f, wrap="word", font=MONO, bg=C["surface"], fg=C["text"], relief="flat",
                    highlightthickness=1, highlightbackground=C["line"], padx=8, pady=8)
        t.tag_configure("err", foreground=C["error"])
        t.tag_configure("head", foreground=C["muted"])
        for when, sheet, text, error in reversed(app.log):
            t.insert("end", f"{when}  {sheet}\n", "head")
            t.insert("end", text.rstrip() + "\n\n", "err" if error else "")
        if not app.log:
            t.insert("end", "No errors or long messages yet.", "head")
        t.configure(state="disabled")
        t.pack(fill="both", expand=True)
        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=(12, 0))
        ttk.Button(bar, text="Close", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="Copy all", command=lambda: (self.clipboard_clear(), self.clipboard_append(
            t.get("1.0", "end").strip()))).pack(side="right", padx=8)
        self.bind("<Escape>", lambda e: self.destroy())


class LookupDialog(tk.Toplevel):
    """Where a column's dropdown comes from: automatic (foreign key or lookup
    file), a table you pick (key column + fields shown beside it), or none;
    and how those fields show in the grid.
    result: "auto", {"off": True}, {"db", "table", "key", "label": [fields]} or None.
    show: "default" / "same" / "separate" / "off" (label display of the column)."""
    SHOW = {"As in Settings": "default", "In the cell: id · value": "same", "In the cell: value only": "value",
            "Own column per field": "separate", "Id only": "off"}

    def __init__(self, app: "App", col: "Col"):
        super().__init__(app.root)
        self.app, self.col, self.result = app, col, None
        self.title(f"Lookup · {col.name}")
        self.configure(bg=C["surface"])
        self.transient(app.root)
        self.resizable(False, False)
        over = app.settings.get("lookups", {}).get(col_key(app.database, app.table, col.name))
        src = col.src or {}
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both")
        self.v_mode = tk.StringVar(value="off" if over and over.get("off") else "table" if over else "auto")
        auto = f"Automatic ({col.lookup})" if (col.lookup and not over) else \
            f"Automatic (foreign key to {col.fk})" if col.fk else "Automatic (none found)"
        for i, (val, text) in enumerate((("auto", auto), ("table", "From a table"), ("off", "No dropdown"))):
            ttk.Radiobutton(f, text=text, value=val, variable=self.v_mode, command=self._mode
                            ).grid(row=i, column=0, columnspan=2, sticky="w", pady=2)
        form = ttk.Frame(f)
        form.grid(row=3, column=0, columnspan=2, sticky="we", padx=(24, 0), pady=(4, 8))
        table = src.get("table", "")
        self.v_db = tk.StringVar(value=src.get("db") or app.database)
        self.v_table, self.v_key = tk.StringVar(value=table), tk.StringVar(value=src.get("key", ""))
        self.chosen = label_fields(src) if src else []  # ticked fields, in tick order
        self.checks: dict[str, tk.BooleanVar] = {}
        self.show = None
        self.widgets = []
        same = self.v_db.get().lower() == app.database.lower()
        for r, (text, var, values) in enumerate((("Database", self.v_db, app.tab.dbs),
                                                 ("Table", self.v_table, app._all_tables if same else []),
                                                 ("Key", self.v_key, []))):
            ttk.Label(form, text=text).grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
            cb = ttk.Combobox(form, textvariable=var, values=values, width=34,
                              state="normal" if r < 2 else "readonly")
            cb.grid(row=r, column=1, sticky="w", pady=3)
            self.widgets.append(cb)
        self.cb_db, self.cb_table, self.cb_key = self.widgets
        self.cb_key.bind("<<ComboboxSelected>>", lambda e: self._fields())
        ttk.Label(form, text="Show").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=(6, 3))
        self.v_find = tk.StringVar()
        self.e_find = ttk.Entry(form, textvariable=self.v_find, width=36)
        self.e_find.grid(row=3, column=1, sticky="w", pady=(6, 3))
        Hint(app, self.e_find, self.v_find, "Find a field")
        self.v_find.trace_add("write", lambda *a: self._fields())
        self.box = ttk.Frame(form)
        self.box.grid(row=4, column=1, sticky="w", pady=(2, 3))
        sub_row = ttk.Frame(form)
        sub_row.grid(row=5, column=1, sticky="we")
        self.l_order = ttk.Label(sub_row, style="Group.TLabel")
        self.l_order.pack(side="left")
        self.b_sub = ttk.Button(sub_row, text="Field lookup…", style="Tool.TButton", command=self._sub)
        self.b_sub.pack(side="right")
        self.subs_changed = False
        own = app.settings.get("label_modes", {}).get(col_key(app.database, app.table, col.name))
        self.v_show = tk.StringVar(value={v: k for k, v in self.SHOW.items()}[own or "default"])
        ttk.Label(form, text="Fields go").grid(row=6, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Combobox(form, textvariable=self.v_show, values=list(self.SHOW), width=34, state="readonly"
                     ).grid(row=6, column=1, sticky="w", pady=3)
        self._fields(col.label_fields if src else [])
        self._tables_of = app.database if same else None
        for ev in ("<<ComboboxSelected>>", "<Return>", "<FocusOut>"):
            self.cb_db.bind(ev, lambda e: self._tables())
        self.cb_table.bind("<<ComboboxSelected>>", lambda e: self._columns())
        self.cb_table.bind("<Return>", lambda e: self._columns())
        self.cb_table.bind("<FocusOut>", lambda e: self._columns())
        ttk.Label(f, text="Any database on this server. Display only: the cell keeps the key.",
                  style="Group.TLabel").grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 12))
        bar = ttk.Frame(f)
        bar.grid(row=5, column=0, columnspan=2, sticky="we")
        ttk.Button(bar, text="View values", command=lambda: app.lookup_view()).pack(side="left")
        ttk.Button(bar, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="Save", style="Primary.TButton", command=self._ok).pack(side="right", padx=8)
        self.bind("<Escape>", lambda e: self.destroy())
        self._loaded = None
        self._mode()
        if not same:
            self._tables()  # then its columns: one job at a time on the connection
        elif table:
            self._columns()
        modal(self)
        self.wait_window()

    def _where(self) -> tuple[str, str]:
        """(database, schema.table); a table typed as database.schema.table wins."""
        parts = [p.strip("[] ") for p in self.v_table.get().strip().split(".")]
        if len(parts) >= 3:
            return parts[0], ".".join(parts[1:3])
        return self.v_db.get().strip() or self.app.database, ".".join(parts)

    def _tables(self):
        """List the tables of the picked database."""
        db = self.v_db.get().strip()
        if not db or db == self._tables_of:
            return
        self._tables_of = db
        dbc = self.app.db

        def done(tables, secs):
            if self.winfo_exists() and self.v_db.get().strip() == db:
                self.cb_table.configure(values=tables)
                if self.v_table.get() and self.v_table.get() not in tables:
                    self.v_table.set("")
                self.app.msg(f"{db}: {len(tables)} tables")
                if self.v_table.get():
                    self._columns()

        self.app.bg("Reading tables", lambda: dbc.tables(db), done)

    def _mode(self):
        mode = self.v_mode.get()
        for w in self.widgets:
            w.state(["!disabled"] if mode == "table" else ["disabled"])
        on = mode == "table" or (mode == "auto" and bool(self.col.src))  # fields of a foreign key too
        for w in (*self.box.winfo_children(), self.e_find, self.b_sub):
            w.state(["!disabled"] if on else ["disabled"])

    SHOWN = 24  # fields listed at once; Find narrows a wide table

    def _fields(self, names: list[str] | None = None):
        """Checklist of the lookup table's columns (all but the key); several can be
        ticked, shown in the order ticked. Ticked fields first, then the rest that
        match Find, at most SHOWN."""
        if names is not None:
            self._names = names
        key = self.v_key.get()
        for w in self.box.winfo_children():
            w.destroy()
        self.checks = {}
        avail = [n for n in self._names if n != key]
        self.chosen = [n for n in self.chosen if n in avail]
        find = self.v_find.get().strip().lower() if hasattr(self, "v_find") else ""
        rest = [n for n in avail if n not in self.chosen and find in n.lower()]
        hidden = max(0, len(self.chosen) + len(rest) - self.SHOWN)
        listed = self.chosen + rest[:max(0, self.SHOWN - len(self.chosen))]
        per_col = max(1, -(-len(listed) // 3))  # up to 3 columns of ticks
        if hidden:
            ttk.Label(self.box, text=f"+{hidden} more: type in Find", style="Group.TLabel"
                      ).grid(row=per_col, column=0, columnspan=3, sticky="w")
        for i, n in enumerate(listed):
            v = self.checks[n] = tk.BooleanVar(value=n in self.chosen)
            ttk.Checkbutton(self.box, text=n, variable=v, command=lambda n=n: self._tick(n)
                            ).grid(row=i % per_col, column=i // per_col, sticky="w", padx=(0, 12))
        if not listed:
            ttk.Label(self.box, text="Pick a table first" if not self._names else
                      "No field matches" if find else "No other columns",
                      style="Group.TLabel").grid(row=0, column=0, sticky="w")
        self._order()
        if hasattr(self, "v_mode"):
            self._mode()

    def _sub(self):
        """Lookup of a ticked field ("SectorID" -> the Sector table): pick the field, then
        the same dialog. Saved now; Save or Cancel here then reloads."""
        if not self.chosen:
            return self.app.msg("Tick a field first", error=True)
        field = self.chosen[0] if len(self.chosen) == 1 else self._pick_field()
        if not field:
            return
        name = f"{self.col.name} ▸ {field}"
        own = self.col.label_fields.index(field) if field in self.col.label_fields else None
        sub = self.col.subs.get(own) if own is not None else None
        sub = sub or Col(name, "", False, False, False, external=True)
        dlg = LookupDialog(self.app, sub)
        if dlg.result is not None:
            self.app.save_lookup(name, dlg.result, dlg.show)
            self.subs_changed = True
            self.app.msg(f"Lookup for {name} saved")
        self.subs_changed |= dlg.subs_changed
        modal(self)

    def _pick_field(self) -> str:
        menu = tk.Menu(self, tearoff=0)
        got = tk.StringVar()
        for f in self.chosen:
            menu.add_command(label=f, command=lambda f=f: got.set(f))
        menu.bind("<Unmap>", lambda e: self.after(50, lambda: got.set(got.get() or "-")))
        menu.tk_popup(self.b_sub.winfo_rootx(), self.b_sub.winfo_rooty() + self.b_sub.winfo_height())
        self.wait_variable(got)
        return "" if got.get() == "-" else got.get()

    def _tick(self, name: str):
        if self.checks[name].get():
            self.chosen.append(name)
        else:
            self.chosen.remove(name)
        self._order()

    def _order(self):
        self.l_order.configure(text=("Shown as: key · " + " · ".join(self.chosen)) if self.chosen
                               else "Nothing ticked: key only")

    def _columns(self):
        """Fill Key / Shown beside from the picked table's columns."""
        db, table = self._where()
        if not table or (db, table) == self._loaded:
            return
        self._loaded = (db, table)
        dbc = self.app.db

        def work():
            return dbc.columns(db, table)

        def done(cols, secs):
            if not self.winfo_exists():
                return
            names = [c.name for c in cols]
            self.cb_key.configure(values=names)
            if self.v_key.get() not in names:
                pk = [c.name for c in cols if c.is_pk]
                self.v_key.set(pk[0] if pk else self.col.name if self.col.name in names else names[0])
            if not any(n in names for n in self.chosen):  # another table: start from the likeliest field
                texts = [c.name for c in cols if c.type in TEXT_TYPES and c.name != self.v_key.get()]
                self.chosen = [g] if (g := guess_label(texts, table)) else []
            self._fields(names)
            self.app.msg(f"{db}.{table}: {len(names)} columns")

        self.app.bg("Reading lookup table", work, done)

    def _ok(self):
        mode, src = self.v_mode.get(), self.col.src
        if mode == "auto" and src and self.chosen != label_fields(src):  # other fields of the same table
            self.result = {**src, "label": list(self.chosen)}
        elif mode == "auto":
            self.result = "auto"
        elif mode == "off":
            self.result = {"off": True}
        else:
            db, table = self._where()
            if not table or not self.v_key.get():
                return self.app.msg("Pick a table and its key column", error=True)
            self.result = {"db": db, "table": table, "key": self.v_key.get(), "label": list(self.chosen)}
        self.show = self.SHOW[self.v_show.get()]
        self.destroy()


class LabelModeDialog(tk.Toplevel):
    """Per-column label display: Settings default, same column, separate, off."""

    def __init__(self, app: "App", col: "Col"):
        super().__init__(app.root)
        self.result = None
        self.title(f"Label display · {col.name}")
        self.configure(bg=C["surface"])
        self.transient(app.root)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both")
        own = app.settings.get("label_modes", {}).get(col_key(app.database, app.table, col.name))
        self.v = tk.StringVar(value=own or "default")
        default = {v: k for k, v in LABEL_MODES.items()}[app.settings.get("label_display", "same")]
        for val, text in (("default", f"As in Settings ({default.lower()})"), *(
                (v, k) for k, v in LABEL_MODES.items())):
            ttk.Radiobutton(f, text=text, value=val, variable=self.v).pack(anchor="w", pady=2)
        ttk.Label(f, text="Display only: the database always gets the id.", style="Group.TLabel"
                  ).pack(anchor="w", pady=(8, 12))
        bar = ttk.Frame(f)
        bar.pack(fill="x")
        ttk.Button(bar, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="Apply", style="Primary.TButton", command=self._ok).pack(side="right", padx=8)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        modal(self)
        self.wait_window()

    def _ok(self):
        self.result = self.v.get()
        self.destroy()


class LookupView(tk.Toplevel):
    """Every allowed value of a lookup column with its label, searchable.
    Use (or double-click) writes the id into the selected cells."""

    def __init__(self, app: "App", col: "Col"):
        super().__init__(app.root)
        self.app, self.col = app, col
        self.title(f"Lookup values · {col.name}")
        self.configure(bg=C["surface"])
        self.transient(app.root)
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both", expand=True)
        self.l_head = ttk.Label(f, text=f"{len(col.choices):,} values from {col.lookup}" if not col.big else
                                f"{col.lookup} · searched on the server, first 200 shown", style="Group.TLabel")
        self.l_head.pack(anchor="w")
        self.v_find = tk.StringVar()
        e = ttk.Entry(f, textvariable=self.v_find)
        e.pack(fill="x", pady=(8, 8))
        Hint(app, e, self.v_find, "Search id or label")
        box = ttk.Frame(f)
        box.pack(fill="both", expand=True)
        fields = col.label_fields or ["label"]
        names = ["id", *(f"f{i}" for i in range(len(fields)))]
        self.tree = ttk.Treeview(box, columns=names, show="headings", selectmode="browse")
        self.tree.heading("id", text="Id", anchor="w")
        self.tree.column("id", width=90, stretch=False)
        for name, field in zip(names[1:], fields):
            self.tree.heading(name, text=field, anchor="w")
            self.tree.column(name, width=max(120, 360 // len(fields)))
        self.geometry(f"{min(1000, 200 + 180 * len(fields))}x560")
        sb = ttk.Scrollbar(box, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.rows = [(k, *col.key_fields.get(k, ("",) * len(fields))) for k in
                     (col.picked.get(e, e) for e in col.choice_list)]
        self._after = None
        self.v_find.trace_add("write", lambda *a: self._search() if col.big else self._fill())
        self.tree.bind("<Double-Button-1>", lambda ev: self._use())
        self.tree.bind("<Return>", lambda ev: self._use())
        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=(12, 0))
        ttk.Button(bar, text="Close", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="Use for selected cells", style="Primary.TButton", command=self._use
                   ).pack(side="right", padx=8)
        self.bind("<Escape>", lambda ev: self.destroy())
        if col.big:
            self._ask()
        else:
            self._fill()
        e.focus_set()

    def _search(self):
        """Big lookup: ask the server once typing pauses."""
        if self._after:
            self.after_cancel(self._after)
        self._after = self.after(300, self._ask)

    def _ask(self):
        self._after = None
        if self.app._busy:  # the connection is in use: try again shortly
            self._after = self.after(300, self._ask)
            return
        text, db, col = self.v_find.get().strip(), self.app.db, self.col

        def done(rows, secs):
            if not self.winfo_exists() or self.v_find.get().strip() != text:
                return
            self.rows = [tuple("" if v is None else str(v) for v in r) for r in rows]
            self.l_head.configure(text=f"{col.lookup} · {len(rows):,} shown"
                                       + (" (first 200: type to narrow)" if len(rows) >= 200 else ""))
            self.app.msg(f"Lookup search: {len(rows):,} rows · {secs:.1f}s")
            self._fill()

        self.app.bg("Searching lookup", lambda: db.big_search(col, text), done)

    def _fill(self):
        if self.col.big:  # already filtered by the server
            self.tree.delete(*self.tree.get_children())
            for row in self.rows:
                if not self.tree.exists(row[0]):
                    self.tree.insert("", "end", iid=row[0], values=row)
            return
        t = self.v_find.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        for row in self.rows:
            if not t or any(t in v.lower() for v in row):
                self.tree.insert("", "end", iid=row[0], values=row)

    def _use(self):
        sel = self.tree.selection()
        if sel:
            if self.col.big:  # known now: no server round trip to label it
                vals = tuple(self.tree.item(sel[0], "values")[1:])
                self.col.key_fields[sel[0]] = vals
                self.col.key_label[sel[0]] = LABEL_SEP.join(v for v in vals if v)
            self.app.put_value(self.col, sel[0])


class ColumnFind(tk.Toplevel):
    """Ctrl+Shift+F: type part of a column name, Enter or double-click selects it."""

    def __init__(self, app: "App"):
        super().__init__(app.root)
        self.app = app
        self.title("Find column")
        self.configure(bg=C["surface"])
        self.transient(app.root)
        self.geometry("360x420")
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        self.v = tk.StringVar()
        e = ttk.Entry(f, textvariable=self.v)
        e.pack(fill="x")
        Hint(app, e, self.v, "Column name")
        self.lb = tk.Listbox(f, activestyle="none", borderwidth=0, highlightthickness=0, font=FONT,
                             bg=C["surface"], fg=C["text"], selectbackground=C["accent_soft"],
                             selectforeground=C["text"])
        self.lb.pack(fill="both", expand=True, pady=(8, 0))
        self.v.trace_add("write", lambda *a: self._fill())
        for w in (e, self.lb):
            w.bind("<Return>", lambda ev: self._go())
            w.bind("<Escape>", lambda ev: self.destroy())
        e.bind("<Down>", lambda ev: (self.lb.focus_set(), self.lb.selection_set(0)))
        self.lb.bind("<Double-Button-1>", lambda ev: self._go())
        self._fill()
        e.focus_set()

    def _fill(self):
        t = self.v.get().strip().lower()
        self.hits = [i for i, c in enumerate(self.app.cols) if t in c.label.lower()]
        self.hits.sort(key=lambda i: not self.app.cols[i].label.lower().startswith(t))
        self.lb.delete(0, "end")
        for i in self.hits:
            self.lb.insert("end", f"{num2alpha(i)}   {self.app.cols[i].label}")
        if self.hits:
            self.lb.selection_set(0)

    def _go(self):
        sel = self.lb.curselection()
        if not self.hits:
            return
        c = self.hits[sel[0] if sel else 0]
        sheet = self.app.sheet
        sheet.select_column(c)
        sheet.see(0, c)
        self.app.l_stats.configure(text=col_info(self.app.cols[c]))
        self.destroy()
        sheet.focus_set()


def modal(win: tk.Toplevel):
    """Grab input for a dialog; another app holding the grab is not fatal."""
    try:
        win.grab_set()
    except tk.TclError:
        pass


def _rgb(color: str) -> tuple[int, int, int]:
    return int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16)


@functools.lru_cache(maxsize=None)
def _coverage(w: int, h: int, r: float) -> tuple[tuple[tuple[float, float], ...], ...]:
    """Per pixel (outer, inner) coverage of a rounded rectangle with a 1px border,
    4x4 supersampled. Colour-free, so every button state and theme reuses it."""
    ss = 4

    def inside(px, py, x0, y0, x1, y1, rr):
        cx, cy = min(max(px, x0 + rr), x1 - rr), min(max(py, y0 + rr), y1 - rr)
        return x0 <= px <= x1 and y0 <= py <= y1 and (px - cx) ** 2 + (py - cy) ** 2 <= rr * rr

    rows = []
    for y in range(h):
        row = []
        for x in range(w):
            outer = inner = 0
            for sy in range(ss):
                for sx in range(ss):
                    px, py = x + (sx + .5) / ss, y + (sy + .5) / ss
                    if inside(px, py, 0, 0, w, h, r):
                        outer += 1
                        inner += inside(px, py, 1, 1, w - 1, h - 1, r - 1)
            row.append((outer / ss ** 2, inner / ss ** 2))
        rows.append(tuple(row))
    return tuple(rows)


def rounded(w: int, h: int, r: float, fill: str, border: str, bg: str) -> tk.PhotoImage:
    """Anti-aliased rounded rectangle with a 1px border, flattened onto `bg`
    (images on Tk widgets have no alpha blending with what is behind them)."""
    f, b, g = _rgb(fill), _rgb(border), _rgb(bg)
    rows = []
    for cov in _coverage(w, h, r):
        rows.append("{" + " ".join(
            "#%02x%02x%02x" % tuple(round(g[k] * (1 - o) + b[k] * (o - i) + f[k] * i) for k in range(3))
            for o, i in cov) + "}")
    img = tk.PhotoImage(width=w, height=h)
    img.put(" ".join(rows))
    return img


class Hint:
    """Muted placeholder text inside an entry or combobox while it is empty."""

    def __init__(self, app: "App", widget, var: tk.StringVar, text: str):
        self.widget, self.var = widget, var
        self.label = app.paint(tk.Label(widget, text=text, font=FONT, cursor="xterm", padx=0, pady=0),
                               bg="surface", fg="muted")
        self.label.bind("<Button-1>", lambda e: widget.focus_set())
        var.trace_add("write", lambda *a: self.update())
        widget.bind("<FocusIn>", lambda e: self.update(), add="+")
        widget.bind("<FocusOut>", lambda e: self.update(), add="+")
        self.update()

    def update(self):
        try:
            typing = self.widget.focus_get() is self.widget
        except (KeyError, tk.TclError):  # focus in a dropdown list
            typing = False
        if self.var.get() or (typing and "readonly" not in self.widget.state()):
            self.label.place_forget()
        else:
            self.label.place(x=9, rely=0.5, anchor="w")


class MenuItem:
    """A menu entry that answers .state([...]) like a ttk button, so it can sit in
    the same enable / disable lists."""

    def __init__(self, menu: tk.Menu, label: str, command):
        menu.add_command(label=label, command=command)
        self.menu, self.index = menu, menu.index("end")

    def state(self, spec):
        self.menu.entryconfigure(self.index, state="disabled" if "disabled" in spec else "normal")


class IconButton(tk.Canvas):
    """32px toolbar icon with a rounded hover fill and a tooltip. `dot` adds a
    status dot (offline grey / online green) at the bottom-right corner."""

    def __init__(self, app: "App", parent, icon: str, command, tip, dot=False):
        super().__init__(parent, width=32, height=32, highlightthickness=0, bd=0, cursor="hand2")
        self.app, self.icon = app, icon
        self.command, self.tip = command, tip
        self.dot: str | None = "offline" if dot else None
        self.hover = False
        self._tipwin = None
        self._fill = None
        self.bind("<Enter>", lambda e: self._set_hover(True))
        self.bind("<Leave>", lambda e: self._set_hover(False))
        self.bind("<Button-1>", lambda e: (self._hide_tip(), self.command()))
        self.redraw()

    def redraw(self):
        self.delete("all")
        self.configure(bg=C["surface"])
        if self.hover:
            self._fill = rounded(32, 32, 6, C["hover"], C["hover"], C["surface"])
            self.create_image(16, 16, image=self._fill)
        self.create_image(16, 16, image=self.app.icon(self.icon))
        if self.dot:
            fill = C["accent"] if self.dot == "online" else C["disabled"]
            self.create_oval(21, 21, 28, 28, fill=fill, outline=C["surface"], width=2)

    def _set_hover(self, on: bool):
        self.hover = on
        self.redraw()
        self._show_tip() if on else self._hide_tip()

    def _show_tip(self):
        text = self.tip() if callable(self.tip) else self.tip
        self._tipwin = tw = tk.Toplevel(self)
        tw.wm_overrideredirect(True)
        tk.Label(tw, text=text, bg=C["surface"], fg=C["text"], font=FONT, padx=8, pady=4,
                 highlightthickness=1, highlightbackground=C["line"]).pack()
        tw.update_idletasks()
        x = self.winfo_rootx() + 32 - tw.winfo_reqwidth()
        tw.wm_geometry(f"+{x}+{self.winfo_rooty() + 36}")

    def _hide_tip(self):
        if self._tipwin:
            self._tipwin.destroy()
            self._tipwin = None


class SettingsDialog(tk.Toplevel):
    """Preferences. Changes apply immediately and are saved."""

    def __init__(self, app: "App"):
        super().__init__(app.root)
        self.title("Settings")
        self.configure(bg=C["surface"])
        self.transient(app.root)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both")
        ttk.Label(f, text="Theme").grid(row=0, column=0, sticky="w", padx=(0, 12), pady=4)
        cb = ttk.Combobox(f, textvariable=app.v_set["theme"], values=["Light", "Dark"],
                          width=10, state="readonly")
        cb.grid(row=0, column=1, sticky="w", pady=4)
        cb.bind("<<ComboboxSelected>>", lambda e: self._theme(app))
        ttk.Checkbutton(f, text="Review before push", variable=app.v_set["review"],
                        command=lambda: app._toggle_setting("review")
                        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Label(f, text="Update shows old → new for every changed cell,\n"
                         "Insert shows the new rows. Nothing is sent until you confirm.",
                  style="Group.TLabel", justify="left").grid(row=2, column=0, columnspan=2,
                                                            sticky="w", pady=(2, 12))
        ttk.Label(f, text="Lookup file").grid(row=3, column=0, sticky="w", padx=(0, 12), pady=4)
        row = ttk.Frame(f)
        row.grid(row=3, column=1, sticky="we", pady=4)
        self.v_lookup = tk.StringVar(value=app.settings.get("lookup_file", ""))
        ttk.Entry(row, textvariable=self.v_lookup, width=34, state="readonly").pack(side="left")
        ttk.Button(row, text="Choose…", command=lambda: self._lookup(app, True)).pack(side="left", padx=(8, 0))
        ttk.Button(row, text="Clear", command=lambda: self._lookup(app, False)).pack(side="left", padx=(4, 0))
        ttk.Label(f, text="JSON of allowed values per column (see README). Foreign keys get a\n"
                         "dropdown on their own. Applies on the next Load or Run.",
                  style="Group.TLabel", justify="left").grid(row=4, column=0, columnspan=2,
                                                            sticky="w", pady=(2, 12))
        ttk.Label(f, text="Lookup labels", style="Group.TLabel").grid(row=8, column=0, columnspan=2, sticky="w")
        for r, (text, key, options) in ((9, ("Show", "label_display", LABEL_MODES)),
                                        (12, ("Not in lookup", "mismatch", MISMATCH))):
            ttk.Label(f, text=text).grid(row=r, column=0, sticky="w", padx=(0, 12), pady=3)
            names = {v: k for k, v in options.items()}
            var = tk.StringVar(value=names.get(app.settings.get(key), next(iter(options))))
            cb = ttk.Combobox(f, textvariable=var, values=list(options), width=22, state="readonly")
            cb.grid(row=r, column=1, sticky="w", pady=3)
            cb.bind("<<ComboboxSelected>>",
                    lambda e, key=key, var=var, options=options: self._label_pref(app, key, options[var.get()]))
        # format and separator: pick one or type your own
        for r, (text, key, values, default) in enumerate((
                ("Format", "label_format", list(LABEL_FORMATS.values()) + ["{label} - {id}", "{id} | {label}"],
                 "{id} · {label}"),
                ("Field separator", "label_sep", LABEL_SEPS, " · ")), start=10):
            ttk.Label(f, text=text).grid(row=r, column=0, sticky="w", padx=(0, 12), pady=3)
            var = tk.StringVar(value=app.settings.get(key, default))
            cb = ttk.Combobox(f, textvariable=var, values=values, width=22)
            cb.grid(row=r, column=1, sticky="w", pady=3)
            for ev in ("<<ComboboxSelected>>", "<Return>", "<FocusOut>"):
                cb.bind(ev, lambda e, key=key, var=var: self._label_text(app, key, var))
        ttk.Label(f, text="Format: type your own with {id} and {label}, e.g. {label} ({id}) or {id} - {label}.\n"
                         "Separator: between several shown fields. Display only: the database always\n"
                         "gets the id. Per column: header menu -> Label display…", style="Group.TLabel",
                  justify="left").grid(row=13, column=0, columnspan=2, sticky="w", pady=(2, 12))
        ttk.Label(f, text="Formula references").grid(row=7, column=0, sticky="w", padx=(0, 12), pady=(0, 12))
        refs = {"Column names (ISIN)": "names", "Cells (B2)": "cells"}
        v_refs = tk.StringVar(value={v: k for k, v in refs.items()}[app.settings.get("formula_refs", "names")])
        cb = ttk.Combobox(f, textvariable=v_refs, values=list(refs), width=22, state="readonly")
        cb.grid(row=7, column=1, sticky="w", pady=(0, 12))
        cb.bind("<<ComboboxSelected>>", lambda e: (app.settings.__setitem__("formula_refs", refs[v_refs.get()]),
                                                   save_settings(app.settings)))
        self.v_ime = tk.BooleanVar(value=app.settings.get("ime", False))
        ttk.Checkbutton(f, text="Input method (IME) for composed characters, applies on restart",
                        variable=self.v_ime, command=lambda: self._ime(app)
                        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 12))
        ttk.Button(f, text="Reset lookups to automatic",
                   command=lambda: (app.settings.pop("fk_labels", None), app.settings.pop("lookups", None),
                                    save_settings(app.settings),
                                    app.msg("Lookups reset to automatic; applies on the next Load"))
                   ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(0, 12))
        ttk.Label(f, text=f"Version {VERSION}", style="Group.TLabel").grid(row=14, column=0, sticky="w")
        ttk.Button(f, text="Close", command=self.destroy).grid(row=14, column=1, sticky="e")
        self.bind("<Escape>", lambda e: self.destroy())
        modal(self)

    def _label_text(self, app: "App", key: str, var: tk.StringVar):
        """A typed format / separator: checked, then applied to every open sheet."""
        value = var.get()
        if key == "label_format":
            err = label_format_error(value)
            if err:
                return app.msg(err, error=True)
        elif not value:
            return
        if app.settings.get(key) == value:
            return
        if key == "label_sep":
            set_label_sep(value)
        self._label_pref(app, key, value)
        app.msg(f"Lookup labels now show as {app._label_fmt('same').format(id='6', label='Valoren' + LABEL_SEP + 'CH')}")

    def _label_pref(self, app: "App", key: str, value: str):
        app.settings[key] = value
        save_settings(app.settings)
        for t in app.tabs:  # every open sheet redraws with the new look
            if t.cols:
                app._in(t, app._relabel)

    def _ime(self, app: "App"):
        app.settings["ime"] = self.v_ime.get()
        save_settings(app.settings)

    def _lookup(self, app: "App", choose: bool):
        path = ""
        if choose:
            path = filedialog.askopenfilename(parent=self, filetypes=[("Lookup JSON", "*.json"), ("All", "*")])
            if not path:
                return
            try:
                load_lookup_file(path)
            except (OSError, ValueError) as e:
                return messagebox.showerror("Lookup file", str(e), parent=self)
        self.v_lookup.set(path)
        app.settings["lookup_file"] = path
        save_settings(app.settings)

    def _theme(self, app: "App"):
        app._toggle_setting("theme")
        self.configure(bg=C["surface"])


class ConnDialog(tk.Toplevel):
    """Saved connections on the left, the selected one's details on the right.
    Passwords are kept in the OS keychain only (see connections.py)."""

    def __init__(self, app: "App"):
        super().__init__(app.root)
        self.app = app
        self.title("Connections")
        self.configure(bg=C["surface"])
        self.transient(app.root)
        self.resizable(False, False)
        self.editing: str | None = app.tab.conn_name

        f = ttk.Frame(self, padding=16)
        f.pack(fill="both")
        ttk.Label(f, text=app.conn_text, foreground=C["accent"] if app.db.cn else C["muted"]
                  ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        left = ttk.Frame(f)
        left.grid(row=1, column=0, sticky="ns", padx=(0, 16))
        ttk.Label(left, text="Saved", style="Group.TLabel").pack(anchor="w", pady=(0, 4))
        self.lb = tk.Listbox(left, height=9, width=24, font=FONT, bg=C["surface"], fg=C["text"],
                             selectbackground=C["accent_soft"], selectforeground=C["text"],
                             highlightthickness=1, highlightbackground=C["line"],
                             highlightcolor=C["accent"], relief="flat", activestyle="none",
                             exportselection=False)
        self.lb.pack(fill="both", expand=True)
        self.lb.bind("<<ListboxSelect>>", lambda e: self._pick())
        self.lb.bind("<Double-Button-1>", lambda e: self._go())
        lb_bar = ttk.Frame(left)
        lb_bar.pack(fill="x", pady=(6, 0))
        ttk.Button(lb_bar, text="New", command=self._new).pack(side="left")
        ttk.Button(lb_bar, text="Remove", command=self._remove).pack(side="left", padx=4)

        form = ttk.Frame(f)
        form.grid(row=1, column=1, sticky="nw")
        self.v_name = tk.StringVar()
        self.v_remember = tk.BooleanVar(value=True)
        rows = (("Name", self.v_name, ""), ("Server", app.v_server, ""), ("User", app.v_uid, ""),
                ("Password", app.v_pwd, "•"))
        for i, (label, var, show) in enumerate(rows):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=4)
            e = ttk.Entry(form, textvariable=var, width=30, show=show)
            e.grid(row=i, column=1, sticky="we", pady=4)
            e.bind("<Return>", lambda ev: self._go())
        keychain = conns.keychain_ok()
        cb = ttk.Checkbutton(form, text="Remember password in OS keychain", variable=self.v_remember)
        cb.grid(row=4, column=1, sticky="w", pady=(4, 0))
        if not keychain:
            self.v_remember.set(False)
            cb.state(["disabled"])
        hint = ("Leave User empty for Windows authentication" if keychain else
                "pip install keyring to remember passwords")
        ttk.Label(form, text=hint, style="Group.TLabel").grid(row=5, column=1, sticky="w", pady=(2, 12))
        bar = ttk.Frame(form)
        bar.grid(row=6, column=0, columnspan=2, sticky="e")
        ttk.Button(bar, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="Connect", style="Primary.TButton", command=self._go
                   ).pack(side="right", padx=8)
        ttk.Button(bar, text="Save", command=self._save).pack(side="right")

        self._refresh(select=self.editing)
        if not self.editing:
            self.v_name.set(self._default_name())
        self.bind("<Escape>", lambda e: self.destroy())
        modal(self)

    def _default_name(self) -> str:
        user, server = self.app.v_uid.get().strip(), self.app.v_server.get().strip()
        return f"{user or 'windows'} @ {server}" if server else ""

    def _refresh(self, select: str | None = None):
        self.lb.delete(0, "end")
        names = [c["name"] for c in conns.all_conns(self.app.settings)]
        for n in names:
            self.lb.insert("end", n)
        if select in names:
            i = names.index(select)
            self.lb.selection_set(i)
            self.lb.see(i)
            self._pick()

    def _pick(self):
        sel = self.lb.curselection()
        if not sel:
            return
        c = conns.find(self.app.settings, self.lb.get(sel[0]))
        self.editing = c["name"]
        self.v_name.set(c["name"])
        self.app.v_server.set(c.get("server", ""))
        self.app.v_uid.set(c.get("user", ""))
        self.app.v_pwd.set(conns.get_password(c["name"]) if c.get("remember") else "")
        self.v_remember.set(bool(c.get("remember")) or not self.app.v_pwd.get())

    def _new(self):
        self.lb.selection_clear(0, "end")
        self.editing = None
        for v in (self.app.v_server, self.app.v_uid, self.app.v_pwd):
            v.set("")
        self.v_name.set("")

    def _save(self) -> bool:
        name = self.v_name.get().strip() or self._default_name()
        if not self.app.v_server.get().strip():
            self.app.msg("Server is required", error=True)
            return False
        self.v_name.set(name)
        warn = conns.save(self.app.settings, name, self.app.v_server.get().strip(),
                          self.app.v_uid.get().strip(), self.app.v_pwd.get(),
                          self.v_remember.get(), old_name=self.editing)
        self.editing = name
        save_settings(self.app.settings)
        self._refresh(select=name)
        self.app.msg(warn or f"Saved connection {name}", error=bool(warn))
        return True

    def _remove(self):
        sel = self.lb.curselection()
        if not sel:
            return
        name = self.lb.get(sel[0])
        if not messagebox.askyesno("Remove connection", f"Remove saved connection '{name}'?\n"
                                   "Its stored password is removed too.", parent=self):
            return
        conns.delete(self.app.settings, name)
        save_settings(self.app.settings)
        self._new()
        self._refresh()

    def _go(self):
        if not self._save():
            return
        self.app.settings["last_conn"] = self.app.tab.conn_name = self.editing
        self.app.tab.tried = True
        self.destroy()
        self.app._save_tabs()
        self.app.connect()


class FilterDialog(tk.Toplevel):
    """Per-column filter: tick values and/or type text the cell must contain.
    Filtering only hides rows; hidden rows still count for Insert / Update."""
    MAX_VALUES = 2000

    def __init__(self, app: "App", c: int):
        super().__init__(app.root)
        self.app, self.c = app, c
        col = app.cols[c]
        self.title(f"Filter · {col.label.rstrip(chr(0x200b))}")
        self.configure(bg=C["surface"])
        self.transient(app.root)
        spec = app.filters.get(col.label, {})
        data = [r for r in app._rows() if not app._spare(r)]
        values = sorted({r[c] or excel_ops.BLANK for r in data},
                        key=lambda v: (v == excel_ops.BLANK, v.lower()))
        self.values = values[:self.MAX_VALUES]

        f = ttk.Frame(self, padding=16)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Contains").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.v_text = tk.StringVar(value=spec.get("text", ""))
        e = ttk.Entry(f, textvariable=self.v_text, width=32)
        e.grid(row=0, column=1, sticky="we", pady=(0, 8))
        e.focus_set()
        self.lb = tk.Listbox(f, selectmode="multiple", height=14, font=FONT, bg=C["surface"],
                             fg=C["text"], selectbackground=C["accent_soft"], selectforeground=C["text"],
                             highlightthickness=1, highlightbackground=C["line"], relief="flat",
                             activestyle="none", exportselection=False)
        self.lb.grid(row=1, column=0, columnspan=2, sticky="nsew")
        chosen = spec.get("values")
        for i, v in enumerate(self.values):
            self.lb.insert("end", v)
            if chosen is None or v in chosen:
                self.lb.selection_set(i)
        more = len(values) - len(self.values)
        note = f"{len(values):,} distinct values" + (f", first {self.MAX_VALUES:,} listed" if more else "")
        ttk.Label(f, text=note, style="Group.TLabel").grid(row=2, column=0, columnspan=2, sticky="w", pady=4)
        sel = ttk.Frame(f)
        sel.grid(row=3, column=0, columnspan=2, sticky="w")
        ttk.Button(sel, text="Select all", command=lambda: self.lb.selection_set(0, "end")).pack(side="left")
        ttk.Button(sel, text="Clear", command=lambda: self.lb.selection_clear(0, "end")).pack(side="left", padx=4)
        bar = ttk.Frame(f)
        bar.grid(row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(bar, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="Apply", style="Primary.TButton", command=self._apply).pack(side="right", padx=8)
        ttk.Button(bar, text="Clear filter", command=self._clear).pack(side="right")
        f.columnconfigure(1, weight=1)
        f.rowconfigure(1, weight=1)
        self.bind("<Return>", lambda ev: self._apply())
        self.bind("<Escape>", lambda ev: self.destroy())
        modal(self)

    def _apply(self):
        picked = {self.values[i] for i in self.lb.curselection()}
        all_listed = len(picked) == len(self.values)  # nothing unticked: no value filter
        spec = {"values": None if all_listed else picked, "text": self.v_text.get()}
        self.destroy()
        self.app.set_filter(self.c, None if spec["values"] is None and not spec["text"].strip() else spec)

    def _clear(self):
        self.destroy()
        self.app.set_filter(self.c, None)


class ReviewDialog(tk.Toplevel):
    """Show what a push will write and ask for confirmation. `fills` maps a column
    index to a palette key used to tint that whole column (old / new)."""

    def __init__(self, parent, title: str, summary: str, headers: list[str], rows: list[list],
                 fills: dict[int, str], action: str):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=C["surface"])
        self.ok = False
        self.transient(parent)
        self.geometry("980x560")
        ttk.Label(self, text=summary, padding=(16, 14, 16, 10), foreground=C["text"]
                  ).pack(anchor="w")
        wrap = tk.Frame(self, bg=C["line"], padx=1, pady=1)
        wrap.pack(fill="both", expand=True, padx=16)
        sh = Sheet(wrap, data=rows, headers=headers, font=(CELL[0], CELL[1], "normal"),
                   header_font=(CELL[0], CELL[1], "normal"), index_font=(CELL[0], CELL[1], "normal"),
                   default_row_height=22, show_x_scrollbar=True, **sheet_colors())
        sh.enable_bindings("single_select", "drag_select", "column_select", "row_select",
                           "copy", "column_width_resize", "arrowkeys", "right_click_popup_menu",
                           "rc_select")
        for ci, key in fills.items():
            sh.highlight_columns([ci], bg=C[key], redraw=False)
        sh.set_all_column_widths()
        sh.pack(fill="both", expand=True)
        bar = ttk.Frame(self, padding=16)
        bar.pack(fill="x")
        ttk.Button(bar, text=action, style="Primary.TButton", command=self._yes).pack(side="right")
        ttk.Button(bar, text="Cancel", command=self.destroy).pack(side="right", padx=8)
        self.bind("<Escape>", lambda e: self.destroy())
        modal(self)
        self.wait_window()

    def _yes(self):
        self.ok = True
        self.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
