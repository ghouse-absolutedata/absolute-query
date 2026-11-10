"""SQL box helpers: suggestions while typing (keywords, tables, columns, joins
from foreign keys) and syntax colours.

The parsing is plain functions (no Tk) so it can be tested on its own. Metadata
comes from a separate read-only connection on a worker thread, so typing never
waits on, or blocks, the sheet's own connection."""
from __future__ import annotations

import queue
import re
import threading
import tkinter as tk

KEYWORDS = [
    "SELECT", "FROM", "WHERE", "AND", "OR", "NOT", "IN", "IS", "NULL", "IS NULL", "IS NOT NULL",
    "LIKE", "BETWEEN", "EXISTS", "AS", "ON", "JOIN", "INNER JOIN", "LEFT JOIN", "RIGHT JOIN",
    "FULL JOIN", "CROSS JOIN", "CROSS APPLY", "OUTER APPLY", "ORDER BY", "GROUP BY", "HAVING",
    "TOP", "DISTINCT", "UNION", "UNION ALL", "EXCEPT", "INTERSECT", "WITH", "CASE", "WHEN",
    "THEN", "ELSE", "END", "ASC", "DESC", "OFFSET", "FETCH NEXT", "ROWS ONLY", "OVER",
    "PARTITION BY", "NOLOCK", "WITH (NOLOCK)",
]
FUNCTIONS = [
    "COUNT", "SUM", "AVG", "MIN", "MAX", "ISNULL", "COALESCE", "NULLIF", "CAST", "CONVERT",
    "TRY_CAST", "TRY_CONVERT", "GETDATE", "SYSDATETIME", "DATEADD", "DATEDIFF", "DATEPART",
    "YEAR", "MONTH", "DAY", "EOMONTH", "FORMAT", "LEN", "LEFT", "RIGHT", "SUBSTRING",
    "CHARINDEX", "REPLACE", "LTRIM", "RTRIM", "TRIM", "UPPER", "LOWER", "CONCAT", "STRING_AGG",
    "ROUND", "ABS", "IIF", "ROW_NUMBER", "RANK", "DENSE_RANK", "LAG", "LEAD", "STUFF",
]
# keywords listed first after a clause (WHERE after FROM, not WHEN)
NEXT_AFTER = {
    c: ("WHERE", "JOIN", "INNER JOIN", "LEFT JOIN", "ORDER BY", "GROUP BY", "WITH (NOLOCK)",
        "CROSS APPLY", "OUTER APPLY") for c in ("FROM", "JOIN", "APPLY")}
NEXT_AFTER["ON"] = ("AND", "WHERE", "JOIN", "INNER JOIN", "LEFT JOIN", "ORDER BY", "GROUP BY")
NEXT_AFTER["WHERE"] = ("AND", "OR", "ORDER BY", "GROUP BY", "IS NULL", "IS NOT NULL", "IN", "LIKE")
# words that end a table reference instead of naming its alias
NOT_ALIAS = {
    "where", "on", "join", "inner", "left", "right", "full", "cross", "outer", "group", "order",
    "with", "union", "except", "intersect", "as", "set", "having", "apply", "pivot", "unpivot",
    "option", "for", "select", "from", "and", "or", "into", "values", "when", "then", "else", "end",
}
_HL_WORDS = {w for k in KEYWORDS for w in k.replace("(", " ").replace(")", " ").split()} | {
    "INTO", "VALUES", "SET", "DECLARE", "BEGIN", "TRAN", "TRANSACTION", "COMMIT", "ROLLBACK",
    "USE", "OUTER", "INNER", "LEFT", "RIGHT", "FULL", "CROSS", "APPLY", "BY", "ALL", "ANY",
    "SOME", "PERCENT", "TIES", "NEXT", "ROWS", "ONLY", "ROW", "FIRST"}

NAME = r"(?:\[(?:[^\]]|\]\])+\]|[A-Za-z_#@][\w#@$]*)"
DOTTED = rf"{NAME}(?:\s*\.\s*{NAME}){{0,2}}"
_MASK = re.compile(r"--[^\n]*|/\*.*?(?:\*/|$)|'(?:[^']|'')*(?:'|$)|\"[^\"]*(?:\"|$)", re.S)
_REF = re.compile(rf"(?:\bFROM|\bJOIN|\bAPPLY|,)\s+({DOTTED})(?:\s+(?:AS\s+)?({NAME}))?", re.I)
_TAIL = re.compile(rf"(?:({DOTTED})\s*\.\s*)?(\[(?:[^\]]|\]\])*|[\w#@$]*)$")
_CLAUSE = re.compile(r"\b(SELECT|FROM|JOIN|APPLY|WHERE|ON|BY|HAVING|SET|INTO|UPDATE|CASE|WHEN|"
                     r"THEN|ELSE|AND|OR|TOP|DISTINCT|UNION|EXCEPT|INTERSECT|VALUES|WITH|OVER)\b", re.I)
_PLAIN = re.compile(r"^[A-Za-z_][\w@#$]*$")


def unq(name: str) -> str:
    name = name.strip()
    return name[1:-1].replace("]]", "]") if name.startswith("[") and name.endswith("]") else name


def quote(name: str) -> str:
    """Bracket a name only when it needs it (spaces, symbols, reserved words)."""
    return name if _PLAIN.match(name) and name.upper() not in _HL_WORDS else "[" + name.replace("]", "]]") + "]"


def split_name(dotted: str) -> list[str]:
    return [unq(p) for p in re.findall(NAME, dotted)]


def mask(sql: str) -> str:
    """Comments and strings blanked out (same length), so words in them do not count."""
    return _MASK.sub(lambda m: m.group(0)[0] + " " * (len(m.group(0)) - 1), sql)


def table_refs(sql: str) -> list[tuple[str, str]]:
    """[(dotted table name as written, alias or "")] of every FROM / JOIN / comma item.
    Comma items of a SELECT list come back too; they are dropped later because they
    name no table."""
    out = []
    for m in _REF.finditer(mask(sql)):
        alias = m.group(2) or ""
        if alias.lower() in NOT_ALIAS:
            alias = ""
        out.append((m.group(1), unq(alias)))
    return out


def context(before: str, full: str) -> dict | None:
    """What fits at the caret. `before` = text up to the caret, `full` = whole text.
    None inside a comment or string. Keys: prefix, qualifier (list of name parts or
    None), clause (last clause keyword, upper case), after (the token right before
    the word: FROM / JOIN / ',' / ...), refs."""
    if _open_at_end(before):
        return None  # caret inside a string or comment
    masked = mask(before)
    m = _TAIL.search(masked)
    prefix = m.group(2)
    qualifier = split_name(m.group(1)) if m.group(1) else None
    head = masked[:m.start()]
    clauses = _CLAUSE.findall(head)
    clause = clauses[-1].upper() if clauses else ""
    prev = re.search(r"([\w\]]+|[,()=<>.*])\s*$", head)
    after = prev.group(1).upper() if prev else ""
    return {"prefix": unq(prefix) if prefix.endswith("]") else prefix.lstrip("["),
            "raw": prefix, "qualifier": qualifier, "clause": clause, "after": after,
            "refs": table_refs(full)}


def _open_at_end(before: str) -> bool:
    """The text ends inside a comment or string that is not closed yet."""
    last = None
    for last in _MASK.finditer(before):
        pass
    if last is None or last.end() != len(before):
        return False
    tok = last.group(0)
    if tok.startswith("--"):
        return True
    if tok.startswith("/*"):
        return not tok.endswith("*/") or len(tok) < 4
    if tok.startswith("'"):
        return not re.fullmatch(r"'(?:[^']|'')*'", tok)
    return not (len(tok) > 1 and tok.endswith('"'))


def initials(name: str, taken: set[str]) -> str:
    """Short alias for a table: Bond_Documents -> bd, IssuerRating -> ir."""
    words = [w for w in re.split(r"[_\s]+|(?<=[a-z])(?=[A-Z])", name) if w]
    base = ("".join(w[0] for w in words) or "t").lower()
    if not _PLAIN.match(base) or base.upper() in _HL_WORDS:
        base = "t" + base
    alias, n = base, 2
    while alias.lower() in taken:
        alias, n = f"{base}{n}", n + 1
    return alias


# ── metadata ────────────────────────────────────────────────────────────────
class Meta:
    """Tables, columns and foreign keys per (connection, database), fetched once on
    a worker thread. `connect(args)` returns an object with sql_objects(db) and
    sql_columns(db, "schema.table"); it is called on the worker thread."""

    def __init__(self, connect):
        self.connect = connect
        self.objects: dict[tuple, list[tuple[str, str, str]]] = {}  # key -> [(schema, name, kind)]
        self.cols: dict[tuple, tuple[list, list]] = {}  # key + (schema.table lower,) -> (cols, fks)
        self.pending: set[tuple] = set()
        self.failed: set[tuple] = set()
        self.jobs: queue.Queue = queue.Queue()
        self.results: queue.Queue = queue.Queue()
        self.conns: dict[tuple, object] = {}
        threading.Thread(target=self._work, daemon=True).start()

    def _work(self):
        while True:
            kind, key, args, database, table = self.jobs.get()
            try:
                cn = self.conns.get(args)
                if cn is None:
                    cn = self.conns[args] = self.connect(args)
                res = cn.sql_objects(database) if kind == "objects" else cn.sql_columns(database, table)
                self.results.put((kind, key, res, None))
            except Exception as e:  # noqa: BLE001 - suggestions just go without it
                self.conns.pop(args, None)
                self.results.put((kind, key, None, e))

    def want_objects(self, args, database) -> list | None:
        key = (args, database.lower())
        if key in self.objects:
            return self.objects[key]
        if key not in self.pending and key not in self.failed:
            self.pending.add(key)
            self.jobs.put(("objects", key, args, database, None))
        return None

    def want_cols(self, args, database, table: str) -> tuple[list, list] | None:
        key = (args, database.lower(), table.lower())
        if key in self.cols:
            return self.cols[key]
        if key not in self.pending and key not in self.failed:
            self.pending.add(key)
            self.jobs.put(("cols", key, args, database, table))
        return None

    def collect(self) -> bool:
        """Take finished fetches (UI thread). True when anything arrived."""
        got = False
        while True:
            try:
                kind, key, res, err = self.results.get_nowait()
            except queue.Empty:
                return got
            got = True
            self.pending.discard(key)
            if err is not None:
                self.failed.add(key)
            elif kind == "objects":
                self.objects[key] = res
            else:
                self.cols[key] = res

    def forget(self):
        """Drop everything cached (Ctrl+Space twice, or after schema changes)."""
        self.objects.clear()
        self.cols.clear()
        self.failed.clear()


# ── suggestions ─────────────────────────────────────────────────────────────
def _score(word: str, prefix: str) -> int | None:
    """0 = starts with the typed text, 1 = a word inside starts with it, 2 = contains
    it, None = no match. Case-insensitive."""
    if not prefix:
        return 0
    w, p = word.lower(), prefix.lower()
    if w.startswith(p):
        return 0
    if re.search(r"(?:^|[_.\s])" + re.escape(p), w) or re.search(r"(?<=[a-z])" + re.escape(prefix), word):
        return 1
    return 2 if p in w else None


def suggest(ctx: dict, objects: list | None, cols_of, fks_of) -> list[tuple[str, str, str]]:
    """[(shown text, text to insert, detail)] for the context, best first.
    cols_of(schema.table) -> [(name, type, is_pk, fk)] or None (not loaded yet);
    fks_of(schema.table) -> [(parent, [(pcol, rcol)], ref)] or None."""
    prefix, qual, clause, after = ctx["prefix"], ctx["qualifier"], ctx["clause"], ctx["after"]
    objects = objects or []
    by_full = {f"{s}.{n}".lower(): (s, n, k) for s, n, k in objects}
    by_name: dict[str, list] = {}
    for s, n, k in objects:
        by_name.setdefault(n.lower(), []).append((s, n, k))

    def resolve(parts: list[str]):
        if len(parts) >= 2:
            return by_full.get(f"{parts[-2]}.{parts[-1]}".lower())
        hits = by_name.get(parts[-1].lower(), [])
        dbo = [h for h in hits if h[0].lower() == "dbo"]
        return (dbo or hits or [None])[0]

    refs = []  # (schema, name, alias)
    for written, alias in ctx["refs"]:
        hit = resolve(split_name(written))
        if hit:
            refs.append((hit[0], hit[1], alias))
    out: list[tuple[int, int, str, str, str]] = []  # (score, group, shown, insert, detail)

    def add(word, insert, detail, group):
        s = _score(word, prefix)
        if s is not None:
            out.append((s, group, word, insert, detail))

    def col_detail(c, alias=""):
        name, typ, pk, fk = c
        bits = [typ] + (["PK"] if pk else []) + ([f"FK → {fk}"] if fk else []) + ([alias] if alias else [])
        return " · ".join(bits)

    tables_here = after in ("FROM", "JOIN", "APPLY", "UPDATE", "INTO") or (after == "," and clause == "FROM")
    if qual:
        # alias. / table. -> its columns; schema. -> its tables
        target = None
        for s, n, a in refs:
            if (len(qual) == 1 and qual[0].lower() in (a.lower(), n.lower())) or \
               (len(qual) >= 2 and [qual[-2].lower(), qual[-1].lower()] == [s.lower(), n.lower()]):
                target = (s, n)
                break
        if target is None and not tables_here:
            hit = resolve(qual)
            target = hit[:2] if hit else None
        if target and not tables_here:
            for c in cols_of(f"{target[0]}.{target[1]}") or []:
                add(c[0], quote(c[0]), col_detail(c), 0)
        else:
            schema = qual[-1].lower()
            for s, n, k in objects:
                if s.lower() == schema:
                    add(n, quote(n), k, 0)
    elif tables_here:
        if after == "JOIN":
            taken = {a.lower() for _, _, a in refs if a} | {n.lower() for _, n, _ in refs}
            for s, n, a in refs:
                for parent, pairs, ref in fks_of(f"{s}.{n}") or []:
                    here = f"{s}.{n}".lower()
                    other = ref if parent.lower() == here else parent
                    os_, on = other.split(".", 1)
                    alias = initials(on, taken)
                    mine = a or quote(n)
                    if parent.lower() == here:  # this table points at `other`
                        cond = " AND ".join(f"{alias}.{quote(rc)} = {mine}.{quote(pc)}" for pc, rc in pairs)
                    else:  # `other` points at this table
                        cond = " AND ".join(f"{alias}.{quote(pc)} = {mine}.{quote(rc)}" for pc, rc in pairs)
                    text = f"{quote(os_)}.{quote(on)} {alias} ON {cond}"
                    add(on, text, "join · foreign key", -1)
        for s, n, k in objects:
            add(n, f"{quote(s)}.{quote(n)}", f"{k} · {s}", 0)
    else:
        column_clause = clause in ("SELECT", "WHERE", "ON", "BY", "HAVING", "SET", "CASE", "WHEN",
                                   "THEN", "ELSE", "AND", "OR", "TOP", "DISTINCT", "OVER")
        if column_clause and refs:
            count: dict[str, int] = {}
            loaded = []
            for s, n, a in refs:
                cs = cols_of(f"{s}.{n}") or []
                loaded.append((s, n, a, cs))
                for c in cs:
                    count[c[0].lower()] = count.get(c[0].lower(), 0) + 1
            for s, n, a, cs in loaded:
                for c in cs:
                    qual_by = a or quote(n)
                    ins = quote(c[0]) if count[c[0].lower()] == 1 or len(refs) == 1 else f"{qual_by}.{quote(c[0])}"
                    add(c[0], ins, col_detail(c, a or n), 0)
        if column_clause:
            for f in FUNCTIONS:
                add(f, f + "(", "function", 2)
        follows = NEXT_AFTER.get(clause, ())
        for k in KEYWORDS:
            add(k, k, "keyword", 0 if k in follows else 1 if not column_clause else 3)
        if not column_clause and prefix:
            for s, n, k in objects:
                add(n, f"{quote(s)}.{quote(n)}", f"{k} · {s}", 4)
    seen, res = set(), []
    for score, group, word, ins, det in sorted(out, key=lambda x: (x[0], x[1], x[2].lower())):
        if (ins, det) not in seen and not (ins.lower() == ctx["raw"].lower() and score == 0 and group >= 0):
            seen.add((ins, det))
            res.append((word if group != -1 else ins, ins, det))
    return res[:300]


# ── Tk side ─────────────────────────────────────────────────────────────────
class SqlAssist:
    """Attach to a tk.Text. `where()` -> (connection args, database) or None when not
    connected; `colors` is the live palette dict; `font` the text font."""

    NAV = {"Up", "Down", "Prior", "Next", "Return", "KP_Enter", "Tab", "Escape", "Home", "End"}

    def __init__(self, text: tk.Text, meta: Meta, where, colors: dict, font):
        self.t, self.meta, self.where, self.C, self.font = text, meta, where, colors, font
        self.items: list[tuple[str, str, str]] = []
        self.ctx: dict | None = None
        self.forced = False
        self.waiting = False  # a list was wanted while metadata was still loading
        self.pop: tk.Toplevel | None = None
        self.lb: tk.Listbox | None = None
        self._hl_after = self._fetch_after = None
        self.recolor()
        t = text
        t.bind("<KeyRelease>", self._key, add="+")
        for k in ("<Up>", "<Down>", "<Prior>", "<Next>", "<Tab>", "<Return>", "<KP_Enter>", "<Escape>"):
            t.bind(k, self._nav)
        t.bind("<Control-space>", self._force)
        t.bind("<FocusOut>", lambda e: t.after(150, self._focus_check), add="+")
        t.bind("<Button-1>", lambda e: self.close(), add="+")
        t.bind("<<Modified>>", self._modified, add="+")
        top = t.winfo_toplevel()
        top.bind("<Configure>", lambda e: self.close() if e.widget is top else None, add="+")
        self._tick()

    # colours ----------------------------------------------------------------
    def recolor(self):
        C = self.C
        self.t.tag_configure("sql_kw", foreground=C["link"])
        self.t.tag_configure("sql_fn", foreground=C["link"])
        self.t.tag_configure("sql_str", foreground=C["accent"])
        self.t.tag_configure("sql_num", foreground=C["text2"])
        self.t.tag_configure("sql_com", foreground=C["muted"])
        if self.pop is not None:
            self.close()

    def _modified(self, _e=None):
        if self.t.edit_modified():
            self.t.edit_modified(False)
            if self._hl_after:
                self.t.after_cancel(self._hl_after)
            self._hl_after = self.t.after(120, self.highlight)
            if self._fetch_after:
                self.t.after_cancel(self._fetch_after)
            self._fetch_after = self.t.after(500, self.prefetch)

    def highlight(self):
        self._hl_after = None
        t = self.t
        sql = t.get("1.0", "end-1c")
        for tag in ("sql_kw", "sql_fn", "sql_str", "sql_num", "sql_com"):
            t.tag_remove(tag, "1.0", "end")
        if len(sql) > 200_000:
            return
        for m in _MASK.finditer(sql):
            tag = "sql_com" if m.group(0)[:1] in "-/" else "sql_str"
            t.tag_add(tag, f"1.0+{m.start()}c", f"1.0+{m.end()}c")
        masked = mask(sql)
        fns = set(FUNCTIONS)
        for m in re.finditer(r"(?<![\w@#$.\[])([A-Za-z_]+)(?![\w@#$])|(?<![\w.])(\d+(?:\.\d+)?)(?![\w])", masked):
            if m.group(1):
                w = m.group(1).upper()
                tag = "sql_kw" if w in _HL_WORDS else "sql_fn" if w in fns and masked[m.end():m.end() + 1] == "(" else None
            else:
                tag = "sql_num"
            if tag:
                t.tag_add(tag, f"1.0+{m.start()}c", f"1.0+{m.end()}c")

    # metadata ---------------------------------------------------------------
    def prefetch(self):
        """Ask for the columns of every table the query names (in the background)."""
        self._fetch_after = None
        w = self.where()
        if not w:
            return
        objs = self.meta.want_objects(*w)
        if objs is None:
            return
        ctx = {"prefix": "", "raw": "", "qualifier": None, "clause": "", "after": "",
               "refs": table_refs(self.t.get("1.0", "end-1c"))}
        for s, n in self._resolved(ctx, objs):
            self.meta.want_cols(w[0], w[1], f"{s}.{n}")

    @staticmethod
    def _resolved(ctx, objs):
        by_full = {f"{s}.{n}".lower(): (s, n) for s, n, _ in objs}
        by_name: dict[str, tuple] = {}
        for s, n, _ in objs:
            if n.lower() not in by_name or s.lower() == "dbo":
                by_name[n.lower()] = (s, n)
        for written, _ in ctx["refs"]:
            p = split_name(written)
            hit = by_full.get(f"{p[-2]}.{p[-1]}".lower()) if len(p) >= 2 else by_name.get(p[-1].lower())
            if hit:
                yield hit

    def _tick(self):
        """Pick up finished metadata fetches; refresh an open list with them."""
        try:
            if self.meta.collect() and (self.pop is not None or self.forced or self.waiting):
                self.waiting = False
                if self._focused() or self.pop is not None:
                    self.update(open_ok=True)
            self.t.after(100, self._tick)
        except tk.TclError:  # widget gone
            pass

    # keys -------------------------------------------------------------------
    def _key(self, e):
        if e.keysym in self.NAV or e.keysym.startswith(("Shift", "Control", "Alt", "Super", "Caps")):
            return
        if e.state & 0x4:  # Ctrl shortcuts (undo, copy, Ctrl+Space itself)
            return
        if e.keysym in ("Left", "Right"):
            return self.update(open_ok=False) if self.pop else None
        ch = e.char
        typed_word = bool(ch) and (ch.isalnum() or ch in "_.[#@$")
        self.update(open_ok=typed_word or e.keysym == "BackSpace" and self.pop is not None
                    or ch == " ")

    def _force(self, _e=None):
        self.forced = True
        self.update(open_ok=True)
        return "break"

    def _nav(self, e):
        if self.pop is None:
            return None
        k = e.keysym
        if k == "Escape":
            self.close()
            return "break"
        if k in ("Return", "KP_Enter", "Tab"):
            if e.state & 0x4:  # Ctrl+Enter runs the query
                self.close()
                return None
            self.accept()
            return "break"
        cur = self.lb.curselection()
        i = cur[0] if cur else 0
        step = {"Up": -1, "Down": 1, "Prior": -10, "Next": 10}[k]
        i = max(0, min(len(self.items) - 1, i + step))
        self.lb.selection_clear(0, "end")
        self.lb.selection_set(i)
        self.lb.see(i)
        return "break"

    def _focused(self) -> bool:
        try:
            return self.t.focus_get() is self.t
        except (KeyError, tk.TclError):  # focus in a dropdown list
            return False

    def _focus_check(self):
        try:
            f = self.t.focus_get()
        except (KeyError, tk.TclError):
            f = None
        if f is not self.t and (self.pop is None or f is None or not str(f).startswith(str(self.pop))):
            self.close()

    # list ---------------------------------------------------------------------
    def update(self, open_ok: bool):
        before = self.t.get("1.0", "insert")
        full = self.t.get("1.0", "end-1c")
        ctx = context(before, full)
        if ctx is None:
            return self.close()
        tables_here = ctx["after"] in ("FROM", "JOIN", "APPLY") or (ctx["after"] == "," and ctx["clause"] == "FROM")
        want = (self.forced or ctx["prefix"] or ctx["raw"] or ctx["qualifier"] is not None
                or (tables_here and before.endswith((" ", "\n", "\t"))))
        if not want or not (open_ok or self.pop is not None):
            return self.close()
        if ctx["prefix"] == "" and not self.forced and ctx["qualifier"] is None and not tables_here:
            return self.close()
        w = self.where()
        objs = self.meta.want_objects(*w) if w else None
        resolved = list(self._resolved(ctx, objs)) if objs else []
        if w:
            for s, n in resolved:
                self.meta.want_cols(w[0], w[1], f"{s}.{n}")
        if w and ctx["qualifier"] and objs:  # a table named only by its qualifier
            p = ctx["qualifier"]
            for s, n, _ in objs:
                if (len(p) >= 2 and [s.lower(), n.lower()] == [p[-2].lower(), p[-1].lower()]) or \
                   (len(p) == 1 and n.lower() == p[0].lower()):
                    self.meta.want_cols(w[0], w[1], f"{s}.{n}")

        def got(table):
            return self.meta.cols.get((w[0], w[1].lower(), table.lower())) if w else None

        self.ctx = ctx
        self.items = suggest(ctx, objs, lambda tb: (got(tb) or (None, None))[0],
                             lambda tb: (got(tb) or (None, None))[1])
        self.forced = False
        if not self.items:
            self.close()
            self.waiting = bool(self.meta.pending)
            return
        self._show()

    def _show(self):
        C = self.C
        if self.pop is None:
            self.pop = tk.Toplevel(self.t)
            self.pop.wm_overrideredirect(True)
            self.pop.configure(bg=C["line"])
            self.lb = tk.Listbox(self.pop, font=self.font, activestyle="none", borderwidth=0,
                                 highlightthickness=0, bg=C["surface"], fg=C["text"], height=10,
                                 selectbackground=C["accent_soft"], selectforeground=C["text"],
                                 exportselection=False)
            sb = tk.Scrollbar(self.pop, command=self.lb.yview)
            self.lb.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            self.lb.pack(side="left", fill="both", expand=True, padx=1, pady=1)
            self.lb.bind("<ButtonRelease-1>", lambda e: self.accept())
        self.lb.delete(0, "end")
        width = min(48, max(len(s) for s, _, _ in self.items))
        for shown, _, det in self.items:
            if len(shown) > width:
                shown = shown[:width - 1] + "…"
            self.lb.insert("end", f"{shown.ljust(width)}  {det}")
        self.lb.configure(width=min(110, width + 2 + max(len(d) for _, _, d in self.items)),
                          height=min(10, len(self.items)))
        self.lb.selection_set(0)
        self.lb.see(0)
        box = self.t.bbox("insert")
        if not box:
            return self.close()
        x, y, _, h = box
        self.pop.geometry(f"+{self.t.winfo_rootx() + x}+{self.t.winfo_rooty() + y + h + 2}")
        self.pop.lift()

    def accept(self):
        if self.pop is None or not self.items:
            return
        cur = self.lb.curselection()
        _, ins, _ = self.items[cur[0] if cur else 0]
        raw = self.ctx["raw"] if self.ctx else ""
        t = self.t
        t.edit_separator()
        if raw:
            t.delete(f"insert-{len(raw)}c", "insert")
        t.insert("insert", ins)
        t.edit_separator()
        self.close()
        t.focus_set()
        self._modified()

    def close(self):
        self.waiting = False
        if self.pop is not None:
            try:
                self.pop.destroy()
            except tk.TclError:
                pass
        self.pop = self.lb = None
        self.items = []
