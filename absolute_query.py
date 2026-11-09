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
  *Description, then *Code); right-click the header -> Lookup label… to pick
  another, show the key only, or turn the dropdown off (remembered per
  referenced table; Settings forgets). Typing a label (or part of it) finds
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
import threading
import tkinter as tk
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
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
    s = {"theme": "Light", "review": True, "lookup_file": "", "ime": False}
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
        self.key_label: dict[str, str] = {}
        self.ref_name = ""  # "database.schema.table" the foreign key points at
        self.label_options: list[str] = []  # its text columns, to show beside the key
        self.label_col = ""  # the one shown
        self.label = name  # grid header; a query alias when it differs from the column name
        self.external = external  # query column not in the target table: shown, never written
        self.readonly = bool(external or is_identity or is_computed
                             or type_ in ("timestamp", "rowversion"))
        self.is_pk = bool(is_pk)
        low = name.lower()
        self.audit = ("cdate" if low in CREATED_DATE else "cby" if low in CREATED_BY else
                      "mdate" if low in MODIFIED_DATE else "mby" if low in MODIFIED_BY else None)


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
        if self.cn is None:
            raise RuntimeError("Not connected")
        setup = ((f"USE {q(database)}; " if database else "") + "SET TRANSACTION ISOLATION LEVEL "
                 f"{'READ COMMITTED' if write else 'READ UNCOMMITTED'};")
        try:
            c = self.cn.cursor()
            c.execute(setup)
            return c
        except pyodbc.Error as e:
            if "08S01" in str(e) or "08003" in str(e):  # dropped link: reconnect once
                self.connect(*self.args)
                c = self.cn.cursor()
                c.execute(setup)
                return c
            raise

    def _read(self, database, sql, params=(), top=None) -> tuple[list[str], list[list[str]]]:
        """One read in its own transaction, committed. Rows arrive in chunks so
        self.fetched shows progress and cancel() can stop it between chunks."""
        c = self.cur(database)
        self.active, self.fetched, self.cancelled = c, 0, False
        try:
            c.execute(sql, *params)
            while c.description is None and c.nextset():
                pass
            if c.description is None:
                raise RuntimeError("Query returned no result set")
            names = [d[0] or f"col{i + 1}" for i, d in enumerate(c.description)]
            rows: list[list[str]] = []
            while top is None or len(rows) < top:
                chunk = c.fetchmany(5000 if top is None else min(5000, top - len(rows)))
                if not chunk:
                    break
                rows += [[v if type(v) is str else fmt(v) for v in r] for r in chunk]  # most cells are text
                self.fetched = len(rows)
            c.close()
            self.cn.commit()
            return names, rows
        except Exception:
            self.cn.rollback()
            if self.cancelled:  # the driver reports a cancel in several ways
                raise RuntimeError("Cancelled") from None
            raise
        finally:
            self.active = None

    def cancel(self):
        """Stop the running statement (called from the UI thread)."""
        c = self.active
        if c is not None:
            self.cancelled = True
            try:
                c.cancel()
            except pyodbc.Error:
                pass

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

    def lookups(self, database: str, table: str, cols: list[Col], spec: dict, labels: dict) -> None:
        """Fill col.choices from the lookup file (`spec`, see lookup_for) or, failing
        that, from the referenced column of a foreign key. A foreign key's dropdown
        shows a text column of the referenced table beside the key: the one chosen
        from the header menu (`labels`, "db.schema.table" -> column, "" = key only,
        NO_LOOKUP = no dropdown), else the likeliest (see guess_label)."""
        jobs = []
        for col in cols:
            if col.external or col.readonly:
                continue
            entry = lookup_for(spec, database, table, col.name)
            if isinstance(entry, list):
                set_choices(col, [str(v) for v in entry], "lookup file")
            elif isinstance(entry, dict) and entry.get("sql"):
                jobs.append((col, entry.get("database") or database, entry["sql"], "lookup file"))
            elif col.fk:
                ref_schema, ref_table, ref_col = col.fk.split(".")
                rows = self._read(database, "SELECT SUM(rows) FROM sys.partitions WHERE object_id = OBJECT_ID(?) "
                                            "AND index_id IN (0, 1)", (f"{q(ref_schema)}.{q(ref_table)}",))[1]
                if rows and rows[0][0] and int(rows[0][0]) > MAX_CHOICES:  # too big to list
                    continue
                col.ref_name = f"{database}.{ref_schema}.{ref_table}"
                label = labels.get(col.ref_name)
                col.label_options = [r[0] for r in self._read(database, """
                    SELECT c.name FROM sys.columns c
                    JOIN sys.types t ON t.user_type_id = c.user_type_id
                    WHERE c.object_id = OBJECT_ID(?) AND c.name <> ?
                      AND t.name IN ('char', 'varchar', 'nchar', 'nvarchar')
                    ORDER BY c.column_id""", (f"{q(ref_schema)}.{q(ref_table)}", ref_col))[1]]
                if label == NO_LOOKUP:  # turned off from the header menu
                    continue
                col.label_col = guess_label(col.label_options) if label is None else label
                label = col.label_col
                jobs.append((col, database, fk_lookup_sql(col.fk, label), f"foreign key to {col.fk}"))
        for col, db, sql, source in jobs:
            err = check_select(sql)
            if err:
                raise RuntimeError(f"Lookup for {col.name}: {err}")
            set_choices(col, self._read(db, sql, top=MAX_CHOICES + 1)[1], source)

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


def clean_err(e: Exception) -> str:
    s = str(e.args[1] if isinstance(e, pyodbc.Error) and len(e.args) > 1 else e)
    if "HY008" in s or "Operation canceled" in s:
        return "Cancelled"
    s = re.sub(r"^\[\w+\]\s*", "", s)
    s = re.sub(r"\[Microsoft\]\[ODBC Driver \d+ for SQL Server\]\[SQL Server\]", "", s)
    return re.sub(r"\s*\(\d+\)\s*\(SQL\w+\)\s*$", "", s).strip()


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
        if not c.readonly and c.audit is None and c.name.lower() not in seen:
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


def cell_error(col: "Col", v: str) -> str | None:
    """Why `v` cannot go into `col`, or None when it can. Blank / NULL always can."""
    if col.external or col.readonly or col.audit:
        return None
    if col.choices is not None and v not in col.choices and v != "" and v.strip().upper() != "NULL":
        return f"Not in the lookup list ({col.lookup})"
    return _value_error(col.type, col.size, col.precision, col.scale, v)


def col_info(col: "Col") -> str:
    """'ISIN · varchar(12) · primary key · lookup: foreign key to dbo.Bond.ISIN'"""
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
    if col.choices is not None:
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


def guess_label(names: list[str]) -> str:
    """The text column that most likely names a lookup row: *Name, then
    *Description, then *Code, then the first; audit columns never."""
    names = [n for n in names if not re.search(r"(created|modified|updated)(by|user)$", n, re.I)]
    for pat in (r"name$", r"desc(ription)?$", r"code$"):
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


def fk_lookup_sql(fk: str, label: str = "") -> str:
    """Values of the column a foreign key references, with `label` beside them."""
    ref_schema, ref_table, ref_col = fk.split(".")
    shown = f", MIN({q(label)})" if label else ""
    return (f"SELECT TOP ({MAX_CHOICES + 1}) {q(ref_col)}{shown} FROM {q(ref_schema)}.{q(ref_table)} "
            f"WHERE {q(ref_col)} IS NOT NULL GROUP BY {q(ref_col)} ORDER BY 1")


def set_choices(col: "Col", rows: list, source: str) -> None:
    """rows: values, or [value, label] pairs (a second column is shown beside the value)."""
    rows = [r if isinstance(r, (list, tuple)) else [r] for r in rows]
    rows = list({str(r[0]): r for r in rows if r and str(r[0]) != ""}.values())
    if len(rows) > MAX_CHOICES:  # too many to list: leave it to the server
        return
    col.choices, col.lookup = frozenset(str(r[0]) for r in rows), source
    col.choice_list, col.picked, col.key_label = [], {}, {}
    for r in rows:
        key, label = str(r[0]), str(r[1]) if len(r) > 1 and r[1] not in (None, "") else ""
        entry = f"{key}  ·  {label}" if label else key
        col.choice_list.append(entry)
        if label:
            col.picked[entry], col.key_label[key] = key, label


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
        self.tabs: list[Tab] = []
        self.tab: Tab | None = None  # tab App code works on; differs from shown only
        self.shown: Tab | None = None  # while a background tab's job finishes (_in)
        self._tab_seq = 0

        root.title(APP_NAME)
        root.geometry("1440x860")
        root.minsize(1100, 520)
        self.settings = load_settings()
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
        self.cb_db.bind("<<ComboboxSelected>>", lambda e: self.pick_db())
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
        self.l_msg = ttk.Label(foot, text="Ready", style="Status.TLabel")
        self.l_msg.pack(side="left", fill="x", expand=True)

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
                header_align="center", index_align="center", **sheet_colors(),
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
        v = self.sheet.get_cell_data(r, c) if r < self.sheet.get_total_rows() else ""
        self.v_fx.set("" if v is None else str(v))
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

    def _validate_edit(self, event):
        """Typed cell edits must fit the column; pasted values are kept and marked
        instead (see _marks), so a big paste never loses data."""
        c = event.column
        if event.eventname == "end_edit_table" and c < len(self.cols):
            col = self.cols[c]
            value = resolve_choice(col, "" if event.value is None else str(event.value))  # label -> key
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
                               ("Lookup label…", self.lookup_label, None)):
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
        mt.bind_class(tag, "<ButtonPress-1>", self._fh_press)
        mt.bind_class(tag, "<B1-Motion>", self._fh_drag)
        mt.bind_class(tag, "<ButtonRelease-1>", self._fh_release)
        s.bind("<<SheetRedrawn>>", lambda e: self._place_handle())

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
            data = self.sheet.get_sheet_data()
            keep = [r for r, row in enumerate(data)
                    if all(excel_ops.matches("" if row[idx[k]] is None else str(row[idx[k]]), spec)
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
                src = [self._cell(r, c) for r in range(r0, r1)]
                vals = ([src[i % len(src)] for i in range(end_r - r1)] if copy
                        else excel_ops.fill_series(src, end_r - r1))
                for i, v in enumerate(vals):
                    writes[(r1 + i, c)] = v
        elif end_c is not None:
            for r in range(r0, r1):
                src = [self._cell(r, c) for c in range(c0, c1)]
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
        tab.job = (label, t0, cancel)
        self._busy = True
        self._refresh_buttons()
        self.msg(label + " …")

        def run():
            try:
                res = work()
                self.jobs.put((tab, done, res, None, t0))
            except Exception as e:  # noqa: BLE001 - surfaced in UI
                self.jobs.put((tab, done, None, e, t0))

        threading.Thread(target=run, daemon=True).start()

    def _poll(self):
        try:
            while True:
                tab, done, res, err, t0 = self.jobs.get_nowait()
                if tab not in self.tabs:
                    continue
                secs = (dt.datetime.now() - t0).total_seconds()

                def finish():
                    self._busy = False
                    self.tab.job = None
                    if err:
                        self._show_error(err)
                    elif done:
                        done(res, secs)
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
        label, t0, _ = t.job
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
        self.shown.db.cancel()

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
        self.tab.note = (text, error)
        if self.tab is self.shown:
            self.l_msg.configure(text=text, style="Err.TLabel" if error else "Status.TLabel")

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
        labels = dict(self.settings.get("fk_labels", {}))

        def work():
            cols = db.columns(database, table)
            db.lookups(database, table, cols, spec, labels)
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
        labels = dict(self.settings.get("fk_labels", {}))

        def work():
            cols = db.columns(database, table)
            db.lookups(database, table, cols, spec, labels)
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
        labels = dict(self.settings.get("fk_labels", {}))
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
                db.lookups(tdb, target, table_cols, spec, labels)
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
        self.filters = {}
        self.sheet.display_rows("all", all_rows_displayed=True)
        self.sheet.set_sheet_data(data, reset_col_positions=True)
        self.sheet.headers([c.label for c in cols] + ["__rid"])
        self.sheet.display_columns(list(range(n)), all_columns_displayed=False)
        self.sheet.readonly_columns([i for i, c in enumerate(cols) if c.readonly or c.audit], True)
        self._dropdowns()
        self._fit_columns(data)
        self._marks()

    def lookup_label(self):
        """Header menu: correct what a foreign key dropdown shows beside the key
        (or turn the dropdown off). Remembered for the referenced table."""
        i = self._col()
        if i is None:
            return
        col = self.cols[i]
        if not col.ref_name:
            return self.msg(f"{col.name} has no foreign key lookup", error=True)
        pick = LabelDialog(self, col).result
        if pick is None:
            return
        self.settings.setdefault("fk_labels", {})[col.ref_name] = pick
        save_settings(self.settings)
        tab, database, db = self.tab, self.database, self.db
        cols = [c for c in self.cols if c.ref_name == col.ref_name]

        def work():
            if pick == NO_LOOKUP:
                return None
            return db._read(database, fk_lookup_sql(col.fk, pick), top=MAX_CHOICES + 1)[1]

        def done(rows, secs):
            for c in cols:
                c.label_col = "" if pick == NO_LOOKUP else pick
                if rows is None:
                    c.choices, c.choice_list, c.picked, c.key_label = None, [], {}, {}
                else:
                    set_choices(c, rows, c.lookup or f"foreign key to {c.fk}")
            if self.tab is tab:
                self._dropdowns()
                self._marks()
                self.msg(f"{col.name}: " + ("no dropdown" if rows is None else
                                            f"showing {pick}" if pick else "key only"))

        self.bg("Reading lookup", work, done)

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
                      if nz(vals[i]) != nz(o[i]) and not self.cols[i].readonly and i not in mirrors]
                if ch:
                    edited.append((gi, rid, ch))
            elif any(nz(v.strip()) is not None for v in vals):
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
        new, edited = self._diff(rows)
        self.sheet.dehighlight_all(redraw=False)
        ro = [i for i, c in enumerate(self.cols) if c.readonly or c.audit]
        if ro:
            self.sheet.highlight_columns(ro, bg=C["readonly"], redraw=False)
        for gi, _ in new:
            self.sheet.highlight_rows([gi], bg=C["new"], redraw=False)
        twin_of = {g[0]: g[1:] for g in self._twins()}
        cells = [(gi, x) for gi, _, ch in edited for ci in ch for x in [ci, *twin_of.get(ci, [])]]
        for i, c in enumerate(self.cols):  # keys and lookups stand out in the header
            if c.is_pk and not c.external:
                self.sheet.highlight_cells(row=0, column=i, canvas="header", fg=C["accent"], redraw=False)
            elif c.choices is not None or c.fk:
                self.sheet.highlight_cells(row=0, column=i, canvas="header", fg=C["link"], redraw=False)
        filtered = [i for i, c in enumerate(self.cols) if c.label in self.filters]
        for i in filtered:
            self.sheet.highlight_cells(row=0, column=i, canvas="header", bg=C["accent_soft"],
                                       fg=C["accent"], redraw=False)
        if cells:
            self.sheet.highlight_cells(cells=cells, bg=C["edited"], redraw=False)
        bad = self.bad = self._bad_cells(rows, new, edited)
        if bad:
            self.sheet.highlight_cells(cells=list(bad), bg=C["old"], fg=C["error"], redraw=False)
        self.sheet.redraw()
        writable = bool(self.table) and any(not c.readonly for c in self.cols)
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
        return bad

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


class LabelDialog(tk.Toplevel):
    """Which column of a referenced table to show beside its key in a dropdown."""
    KEY, NONE = "(key only)", "(no dropdown)"

    def __init__(self, app: "App", col: "Col"):
        super().__init__(app.root)
        self.result: str | None = None
        self.title("Lookup label")
        self.configure(bg=C["surface"])
        self.transient(app.root)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both")
        ttk.Label(f, text=f"{col.name} refers to {col.fk}.\n"
                         f"Which column should the dropdown show beside the key?",
                  justify="left").pack(anchor="w", pady=(0, 8))
        current = (self.NONE if col.choices is None else col.label_col or self.KEY)
        self.v = tk.StringVar(value=current)
        cb = ttk.Combobox(f, textvariable=self.v, values=[*col.label_options, self.KEY, self.NONE],
                          state="readonly", width=36)
        cb.pack(anchor="w")
        ttk.Label(f, text="Remembered for this table. Settings can reset it.",
                  style="Group.TLabel").pack(anchor="w", pady=(4, 12))
        bar = ttk.Frame(f)
        bar.pack(fill="x")
        ttk.Button(bar, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="OK", style="Primary.TButton", command=self._ok).pack(side="right", padx=8)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        modal(self)
        cb.focus_set()
        self.wait_window()

    def _ok(self):
        v = self.v.get()
        self.result = {self.KEY: "", self.NONE: NO_LOOKUP}.get(v, v)
        self.destroy()


class _NoTooltip(tk.Toplevel):
    """Stand-in for tksheet's cell / header / index tooltip windows. tksheet
    builds three per sheet up front and laying them out costs about a second
    per sheet; this app never turns tksheet tooltips or notes on."""
    cell_readonly = note_readonly = True

    def __init__(self, parent, **_):
        super().__init__(parent)
        self.withdraw()
        self.overrideredirect(True)

    def reset(self, *a, **k):
        pass


for _m in (_tks_mt, _tks_ch, _tks_ri):
    _m.Tooltip = _NoTooltip


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
        self.v_ime = tk.BooleanVar(value=app.settings.get("ime", False))
        ttk.Checkbutton(f, text="Input method (IME) for composed characters, applies on restart",
                        variable=self.v_ime, command=lambda: self._ime(app)
                        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 12))
        ttk.Button(f, text="Forget lookup label choices",
                   command=lambda: (app.settings.pop("fk_labels", None), save_settings(app.settings),
                                    app.msg("Lookup labels reset; applies on the next Load"))
                   ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(0, 12))
        ttk.Label(f, text=f"Version {VERSION}", style="Group.TLabel").grid(row=7, column=0, sticky="w")
        ttk.Button(f, text="Close", command=self.destroy).grid(row=7, column=1, sticky="e")
        self.bind("<Escape>", lambda e: self.destroy())
        modal(self)

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
        data = app.sheet.get_sheet_data()
        values = sorted({("" if r[c] is None else str(r[c])) or excel_ops.BLANK for r in data},
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
