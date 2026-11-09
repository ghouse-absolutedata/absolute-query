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
