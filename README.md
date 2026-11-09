# AbsoluteQuery

Spreadsheet-style editor for SQL Server tables: load or query rows, paste from
Excel / CSV, edit, then insert or update in one all-or-nothing transaction.
Windows and Linux.

## Download

| | |
|---|---|
| Windows | [AbsoluteQuery.exe](https://github.com/ghouse-absolutedata/absolute-query/releases/latest/download/AbsoluteQuery.exe) |
| Linux | [AbsoluteQuery-linux](https://github.com/ghouse-absolutedata/absolute-query/releases/latest/download/AbsoluteQuery-linux) |

1. Install the **Microsoft ODBC Driver 18 for SQL Server**:
   - Windows: <https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server>
   - Linux: the `msodbcsql18` package from Microsoft's repository (it brings `unixODBC`).
2. Put the program somewhere you can write to, such as Documents or the desktop on
   Windows, or `~/bin` on Linux. It replaces itself when it updates.
3. Start it:
   - Windows: open `AbsoluteQuery.exe`. The first time, Windows may warn about an
     unknown publisher; choose *More info → Run anyway*.
   - Linux: `chmod +x AbsoluteQuery-linux && ./AbsoluteQuery-linux`
4. Click the server icon (top right) to add a connection.

## Using it

- **Load** a table, with an optional *Where* and *Top*. Paste a column copied from
  a spreadsheet (header line, then values) into *Where* to get `Header IN (...)`.
- **SQL** runs any single `SELECT` / `WITH`. You don't need to pick a database or
  table first. The result can be edited in the table its columns come from.
  Joined and computed columns are read-only.
- Reads never lock: every read runs `READ UNCOMMITTED`, the same as `NOLOCK`.
  Long reads show rows read so far and can be cancelled.
- Values are checked against the column type. Typed bad values are refused.
  Pasted bad values are marked red, and Insert / Update wait until they are fixed.
- Key columns have accent headers. Columns with allowed values have blue headers
  and a dropdown; hover a header for its type and key details.
- Lookup labels: Settings → *Lookup labels* sets how an id shows with its label.
  *Format* takes your own text with `{id}` and `{label}` (`{id} - {label}`,
  `{label} ({id})`, `{label}`); *Field separator* goes between several shown fields
  (` · `, ` - `, ` | ` or your own). Display only: the database gets the id.
- Scratch columns: right-click a header → *Add scratch column* for a column of your
  own, for notes or working things out with formulas. It stays on the sheet until the
  next Load and is never written: Insert / Update skip it, and a row with only a
  scratch value is not a new row. *Rename…* / *Remove* from the same menu.
- **Ctrl+F** finds a value, including the lookup label shown with an id.
  **Ctrl+Shift+F** jumps to a column by name.

### Formulas

Start a cell with `=`, as in Excel. The rest is a SQL Server expression, and it can
use other cells. The cell shows the result and keeps the formula: select it to see
the formula in the fx bar, press F2 to edit it. When a cell it uses changes, it is
calculated again. Insert / Update write the result.

| Type in a cell | Means |
|---|---|
| `=UPPER(ISIN)` | ISIN of the same row, in capitals |
| `=[Issue Date]` | a column whose name has spaces or is a SQL word |
| `=UPPER(B2)` | the cell in column B, row 2 |
| `=$B$2 * 2` | `$` keeps the column / row fixed when copied |
| `=SUM(C2:C10)` | `SUM`, `MIN`, `MAX`, `AVG` / `AVERAGE`, `COUNT` over a range |
| `=DATEADD(month, 1, IssueDate)` | any SQL Server function |
| `=GETDATE()`, `=NEWID()` | no cell needed; `NEWID()` differs in every cell |
| `=TRIM('  QC  9 ')` | text goes in single quotes |
| `=(SELECT MAX(EntityID) FROM FI_EntityMaster.dbo.EntityMaster)` | a read-only subquery |

- While typing a formula, click a cell to put its reference in; drag for a range;
  after an operator, the arrow keys point at cells too. Settings → *Formula
  references* picks what a click on the same row puts in: the column name
  (`ISIN`, the default) or the cell (`A2`). Other rows always give `A2`.
- Fill down / right (Ctrl+D, Ctrl+R, the fill handle) and copy / paste move cell
  references, as in Excel: `=UPPER(B2)` filled down becomes `=UPPER(B3)`. `$B$2`
  stays, and column names always mean the row the formula is in.
- Text needs quotes: `=TRIM('QC  9')`. Without them `QC` is read as a column name.
- Referenced values are read as their column's type, and the result is converted
  to the type of the cell's column, so `=GETDATE()` fits a date column.
- Inside a subquery, name the table's columns with the table
  (`t.EntityID`): a bare name that matches a column of the sheet means that cell.
- A formula that cannot be worked out keeps its cell red; hover it for why
  (server error, a missing cell, a circular reference). Insert / Update wait until
  every formula has a result.
- Read-only: a formula with `INSERT`, `UPDATE`, `DELETE`, `DROP`, `;` and the like
  is refused.
- Typing a value over a formula cell replaces the formula. Formulas last until the
  sheet is loaded again; a value stored in the table that starts with `=` is never
  calculated.
- Cells with the same formula shape (one filled down) are calculated in one query,
  so thousands of rows take about as long as one.

### Lookup file

Foreign keys get a dropdown on their own, with a text column of the referenced
table shown beside each key (`3 · ISIN`). The cell keeps only the key. The label
column is picked automatically: a `…Name` column first, then `…Description`, then
`…Code`.

Right-click any column header and choose **Lookup…** to set or correct a lookup:
- **From a table**: pick any table (`database.schema.table` works too), its key
  column, and the column shown beside the key. This works for columns with no
  foreign key as well.
- **No dropdown**.
- **Automatic**: back to the foreign key or the lookup file.

The choice is remembered per column. *Reset lookups to automatic* in Settings
clears all of them. **View lookup values…** in the same menu lists every id with
its label and has a search box. *Use for selected cells* writes the chosen id into
the selected cells.

You can type a key, a label, or part of a label in the cell or the fx bar. The
dropdown jumps to the matching entry, and Enter stores its key: `valoren` becomes
`6`.

For other columns, choose a JSON lookup file in Settings. A `SELECT` with two
columns shows the second beside the first in the same way:

```json
{
  "Sales.dbo.Orders.Status": ["Open", "Shipped", "Closed"],
  "dbo.Orders.Currency": {"sql": "SELECT Code FROM dbo.Currency", "database": "Ref"},
  "CountryCode": {"sql": "SELECT Code FROM dbo.Country"}
}
```

A key names a column as `database.schema.table.column`, `schema.table.column`,
`table.column`, or just `column` for any table. The most specific key wins. A
value is either the list of allowed values or a `SELECT` returning them. Only
those values, or blank for NULL, are accepted. Edits to the file apply on the
next Load or Run.

**Updates are automatic.** Each start checks this repository and runs the newest
code. There is nothing to clone or reinstall. Without network access, the last
downloaded version runs. Settings (gear icon) shows the running version.

Where your files live:

| | Windows | Linux |
|---|---|---|
| Downloaded code, dry-run CSVs, `launcher.log` | `%LOCALAPPDATA%\AbsoluteQuery` | `~/.local/share/AbsoluteQuery` |
| Settings, open sheets, saved connections (no passwords) | `%USERPROFILE%\.absolutequery.json` | `~/.absolutequery.json` |
| Passwords | Windows Credential Manager | Secret Service keyring |

## How updates work

The program you download is `launcher.py` packaged with Python and every
package the app uses. On each start it:

1. **Checks for a new program.** It compares its build id with the one on the
   `build` release. The build id is a hash of `launcher.py`, `build.py` and
   `requirements.txt`. If they differ, it downloads the new program, swaps it
   in and restarts.
   - This is rare: it happens only when one of those three files changes.
   - Each new build is tried once, so a bad release cannot cause a restart loop.
2. **Checks for new code.** It reads the latest commit on `main` from git's ref
   listing. This avoids GitHub's limit of 60 API calls per hour per IP address.
   - If the commit is new, it downloads that commit's zip archive and unpacks it
     into `versions/<commit>`. It keeps the current and the previous version.
3. **Runs the app** from the newest version that imports cleanly. The order is:
   current, then previous, then the copy bundled inside the program. Reasons
   for any fallback go to `launcher.log`.

A push to `main` therefore reaches every user on their next start.

On every push, GitHub Actions (`.github/workflows/build.yml`):
- builds `AbsoluteQuery.exe` on Windows;
- builds `AbsoluteQuery-linux` on Debian 11, whose glibc 2.31 runs on most
  distros from 2020 on;
- replaces the `build` release with the new files.

Two things to keep in mind:
- **Keep `main` releasable.** Do unfinished work on other branches.
- **New imports get bundled automatically.** When the app starts importing a
  package that is not in `requirements.txt`, add it there. That changes the build
  id, so every user's program picks up the new package by itself.

## Development

```
pip install -r requirements.txt
python absolute_query.py          # run from source
pip install pyinstaller
python build.py                   # local build -> dist/
```

To try a launcher against a fork or branch, set the `ABSOLUTEQUERY_REPO` and
`ABSOLUTEQUERY_BRANCH` environment variables.

Default credentials for the first connection can come from the environment, or
from a `.env` file next to the app: `DB_SERVER`, `DB_UID`, `DB_PWD`. Never commit
a `.env` file; it is in `.gitignore`.

| File | |
|---|---|
| `absolute_query.py` | the app (usage notes in its docstring) |
| `connections.py` | saved connections, passwords in the OS keychain |
| `excel_ops.py` | fill-series and filter helpers |
| `launcher.py` | the downloaded program: update checks, then runs the app |
| `build.py` | PyInstaller build of the launcher |
| `.github/workflows/build.yml` | CI build and release |
