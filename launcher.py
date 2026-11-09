"""AbsoluteQuery launcher - the program users download (built by CI with build.py).

On every start:
1. Program update: if the GitHub release has a program built from a different
   launcher / package set than this one, it is downloaded, swapped in and
   restarted. Rare - only when launcher.py, build.py or requirements.txt change.
2. Code update: the latest commit on BRANCH is looked up; when it differs from
   the cached copy, the repo archive is unpacked into the user's data folder
   and the app runs from there. So a push reaches every user on their next start.

Offline, GitHub down, or an update that fails to import: the newest copy that
works is used - current, then previous, then the copy bundled in the program.

REPO_URL / BRANCH can be overridden at run time with the ABSOLUTEQUERY_REPO /
ABSOLUTEQUERY_BRANCH environment variables.
"""
from __future__ import annotations

import importlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import traceback
import urllib.request
import zipfile
from pathlib import Path

REPO_URL = "https://github.com/ghouse-absolutedata/absolute-query"
BRANCH = "main"
RELEASE_TAG = "build"  # rolling release the CI workflow replaces on every push
ASSET = "AbsoluteQuery.exe" if os.name == "nt" else "AbsoluteQuery-linux"
TIMEOUT = 5  # seconds per check; a slow or missing server never blocks start-up for long
DOWNLOAD_TIMEOUT = 120
ENTRY = "absolute_query"  # module whose main() is the app
APP_MODULES = ("absolute_query", "connections", "excel_ops")
KEEP = 2  # unpacked versions kept: current + previous


def data_dir() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    d = base / "AbsoluteQuery"
    d.mkdir(parents=True, exist_ok=True)
    return d


DATA = data_dir()
VERSIONS = DATA / "versions"
STATE = DATA / "state.json"
LOG = DATA / "launcher.log"


def log(msg: str) -> None:
    try:
        if LOG.is_file() and LOG.stat().st_size > 1_000_000:
            LOG.unlink()
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(msg.rstrip() + "\n")
    except OSError:
        pass


def read_state() -> dict:
    try:
        return json.loads(STATE.read_text())
    except (OSError, ValueError):
        return {}


def write_state(s: dict) -> None:
    tmp = STATE.with_suffix(".tmp")
    tmp.write_text(json.dumps(s))
    os.replace(tmp, STATE)


def get(url: str, timeout: float = TIMEOUT) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "AbsoluteQuery-launcher"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def latest_sha(repo: str, branch: str) -> str:
    """Commit id at the tip of `branch`. Read from git's own ref listing rather than
    the GitHub API, which allows only 60 anonymous calls an hour per IP address -
    an office behind one IP would run out."""
    refs = get(f"{repo.rstrip('/')}.git/info/refs?service=git-upload-pack").decode("latin-1")
    m = re.search(rf"([0-9a-f]{{40}}) refs/heads/{re.escape(branch)}(?:\x00|\n|$)", refs)
    if not m:
        raise RuntimeError(f"branch {branch} not found")
    return m.group(1)


def download(repo: str, sha: str) -> Path:
    """Unpack the repo at `sha` into versions/<sha>; returns the folder holding ENTRY.py."""
    dest = VERSIONS / sha[:12]
    if (dest / f"{ENTRY}.py").is_file():
        return dest
    blob = get(f"{repo.rstrip('/')}/archive/{sha}.zip", DOWNLOAD_TIMEOUT)
    tmp = VERSIONS / f".{sha[:12]}.part"
    shutil.rmtree(tmp, ignore_errors=True)
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        entry = next((n for n in z.namelist() if n.endswith(f"{ENTRY}.py")
                      and n.count("/") <= 2), None)
        if entry is None:
            raise RuntimeError(f"{ENTRY}.py not found in the repo archive")
        root = entry[: -len(f"{ENTRY}.py")]  # the folder in the archive that holds the app
        for n in z.namelist():
            if n.startswith(root) and not n.endswith("/"):
                out = tmp / n[len(root):]
                if ".." in Path(n[len(root):]).parts:
                    continue
                out.parent.mkdir(parents=True, exist_ok=True)
                out.write_bytes(z.read(n))
    shutil.rmtree(dest, ignore_errors=True)
    tmp.rename(dest)
    return dest


def build_id() -> str:
    """Id of this program's build (see build.py); '' when running from source."""
    try:
        return (Path(sys._MEIPASS) / "build_id.txt").read_text().strip()
    except (AttributeError, OSError):
        return ""


def update_program(state: dict) -> None:
    """Replace this program with the release's when that was built differently,
    then restart. Each new build is tried once, so a bad release cannot loop."""
    mine = build_id()
    if not mine:
        return
    exe = Path(sys.executable)
    for old in exe.parent.glob(f"{exe.stem}.old*"):  # left by a previous swap on Windows
        try:
            old.unlink()
        except OSError:
            pass
    repo = os.environ.get("ABSOLUTEQUERY_REPO", REPO_URL).rstrip("/")
    base = f"{repo}/releases/download/{RELEASE_TAG}"
    try:
        latest = get(f"{base}/{ASSET}.build").decode().strip()
        if not latest or latest == mine or state.get("program_tried") == latest:
            return
        state["program_tried"] = latest
        write_state(state)
        blob = get(f"{base}/{ASSET}", DOWNLOAD_TIMEOUT)
        if len(blob) < 1_000_000:
            raise RuntimeError(f"download too small ({len(blob)} bytes)")
        new = exe.with_name(exe.name + ".new")
        new.write_bytes(blob)
        new.chmod(0o755)
        if os.name == "nt":  # a running exe cannot be overwritten, but it can be renamed
            os.replace(exe, exe.with_name(f"{exe.stem}.old{exe.suffix}"))
        os.replace(new, exe)
        log(f"program updated {mine} -> {latest}, restarting")
    except Exception as e:  # noqa: BLE001 - offline, no release yet, read-only folder: carry on
        log(f"program update skipped: {e!r}")
        return
    # fresh start of the new program (not inside this one's unpacked files)
    subprocess.Popen([str(exe), *sys.argv[1:]], env={**os.environ, "PYINSTALLER_RESET_ENVIRONMENT": "1"})
    sys.exit(0)


def update(state: dict) -> None:
    repo = os.environ.get("ABSOLUTEQUERY_REPO", REPO_URL)
    branch = os.environ.get("ABSOLUTEQUERY_BRANCH", BRANCH)
    try:
        sha = latest_sha(repo, branch)
        if sha == state.get("sha") and Path(state.get("dir", "")).is_dir():
            return
        d = download(repo, sha)
        state["previous"] = state.get("dir")
        state.update(sha=sha, dir=str(d))
        write_state(state)
        log(f"updated to {sha[:12]}")
        prune(state)
    except Exception as e:  # noqa: BLE001 - offline / server down: run what we have
        log(f"update check failed: {e!r}")


def prune(state: dict) -> None:
    keep = {Path(p).name for p in (state.get("dir"), state.get("previous")) if p}
    for d in sorted(VERSIONS.glob("*"), key=lambda p: p.stat().st_mtime, reverse=True):
        if d.is_dir() and d.name not in keep and not d.name.startswith("."):
            shutil.rmtree(d, ignore_errors=True)


def bundled() -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / "app" if (base / "app").is_dir() else base


def load(folder: Path):
    """Import the app from `folder`; any earlier attempt's modules are dropped first."""
    for m in APP_MODULES:
        sys.modules.pop(m, None)
    sys.path.insert(0, str(folder))
    try:
        return importlib.import_module(ENTRY)
    except BaseException:
        sys.path.remove(str(folder))
        raise


def show_error(text: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("AbsoluteQuery", text)
        root.destroy()
    except Exception:  # noqa: BLE001
        print(text, file=sys.stderr)


def main() -> None:
    VERSIONS.mkdir(exist_ok=True)
    state = read_state()
    update_program(state)
    update(state)
    tries = [(Path(d).name, d) for d in (state.get("dir"), state.get("previous")) if d]
    for label, folder in tries + [("bundled", str(bundled()))]:
        if not folder or not (Path(folder) / f"{ENTRY}.py").is_file():
            continue
        os.environ["ABSOLUTEQUERY_VERSION"] = label  # read by the app at import
        os.environ["ABSOLUTEQUERY_DATA"] = str(DATA)
        try:
            app = load(Path(folder))
        except Exception:  # noqa: BLE001 - a broken update falls back to the next copy
            log(f"cannot import {label} ({folder}):\n{traceback.format_exc()}")
            continue
        try:
            app.main()
        except Exception:  # noqa: BLE001
            log(traceback.format_exc())
            show_error(f"AbsoluteQuery stopped with an error.\n\nDetails: {LOG}")
        return
    show_error(f"AbsoluteQuery could not start.\n\nDetails: {LOG}")


if __name__ == "__main__":
    main()
