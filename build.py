"""Build the installable AbsoluteQuery program with PyInstaller.

    pip install -r requirements.txt pyinstaller
    python build.py            # -> dist/AbsoluteQuery.exe (Windows) or dist/AbsoluteQuery (Linux)

The GitHub workflow (.github/workflows/build.yml) runs this on Windows and Linux
for every push and publishes the results; running it by hand is only needed for
local testing.

The program carries a build id: a hash of the files that shape the program itself
(launcher, build script, requirements). The launcher compares it with the
release's to decide whether the program needs replacing; app code changes alone
never do - they arrive through the launcher's code update.

Every module the app imports (found by reading its imports) is bundled, plus a
handful of common standard-library modules, so later code changes rarely need a
rebuild. The current app files are bundled too, as the offline fallback.
"""
from __future__ import annotations

import ast
import hashlib
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
APP_FILES = ["absolute_query.py", "connections.py", "excel_ops.py", "icons.py", "sql_assist.py"]
PROGRAM_FILES = ["launcher.py", "build.py", "requirements.txt"]  # define the build id
LOCAL = {Path(f).stem for f in APP_FILES}
# packages pulled in whole (plugins / backends are loaded dynamically)
COLLECT = ["tksheet", "keyring", "certifi"]
# insurance for code pushed later: cheap standard-library modules an update might start using
EXTRA = ["sqlite3", "uuid", "hashlib", "urllib.request", "urllib.parse", "zipfile", "statistics",
         "difflib", "textwrap", "shlex", "subprocess", "webbrowser", "xml.etree.ElementTree",
         "tkinter.simpledialog", "tkinter.scrolledtext", "tkinter.colorchooser", "concurrent.futures",
         "openpyxl", "secretstorage", "jeepney", "win32ctypes"]


def app_imports() -> set[str]:
    mods = set()
    for f in APP_FILES:
        for node in ast.walk(ast.parse((HERE / f).read_text(encoding="utf-8"))):
            if isinstance(node, ast.Import):
                mods.update(a.name for a in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module and not node.level:
                mods.add(node.module)
                mods.update(f"{node.module}.{a.name}" for a in node.names)
    return {m for m in mods if m.split(".")[0] not in LOCAL | {"__future__"}}


def available(mod: str) -> bool:
    try:
        __import__(mod)
        return True
    except Exception:  # noqa: BLE001 - e.g. win32ctypes on Linux
        return False


def build_id() -> str:
    h = hashlib.sha256()
    for f in PROGRAM_FILES:
        h.update((HERE / f).read_bytes().replace(b"\r\n", b"\n"))
    return h.hexdigest()[:16]


def main():
    bid = build_id()
    (HERE / "build").mkdir(exist_ok=True)
    (HERE / "build" / "build_id.txt").write_text(bid)
    hidden = sorted(m for m in app_imports() | set(EXTRA) if available(m))
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onefile", "--windowed",
           "--name", "AbsoluteQuery", "--distpath", str(HERE / "dist"),
           "--workpath", str(HERE / "build"), "--specpath", str(HERE / "build")]
    for f in APP_FILES:
        cmd += ["--add-data", f"{HERE / f}{os.pathsep}app"]
    cmd += ["--add-data", f"{HERE / 'build' / 'build_id.txt'}{os.pathsep}."]
    for m in hidden:
        cmd += ["--hidden-import", m]
    for p in COLLECT:
        cmd += ["--collect-all", p]
    icon = HERE / ("icon.ico" if os.name == "nt" else "icon.png")
    if icon.is_file():
        cmd += ["--icon", str(icon)]
    cmd.append(str(HERE / "launcher.py"))
    print("bundling:", ", ".join(hidden))
    subprocess.run(cmd, check=True)
    (HERE / "dist" / "build_id.txt").write_text(bid)
    print("built:", next((HERE / "dist").glob("AbsoluteQuery*")), "build id", bid)


if __name__ == "__main__":
    main()
