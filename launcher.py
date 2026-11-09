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
import ssl
import subprocess
import sys
import threading
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
APP_MODULES = ("absolute_query", "connections", "excel_ops", "icons")
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


def tls_context() -> ssl.SSLContext:
    """System CAs plus certifi's. The Linux program is built on Debian and its
    OpenSSL looks for Debian's CA path, which other distros do not have."""
    ctx = ssl.create_default_context()
    try:
        import certifi
        ctx.load_verify_locations(certifi.where())
    except Exception:  # noqa: BLE001 - running from source without certifi
        pass
    return ctx


TLS = tls_context()


def get(url: str, timeout: float = TIMEOUT, progress=None) -> bytes:
    """GET `url`. `progress(done, total)` is called as it arrives (total 0 = unknown)."""
    req = urllib.request.Request(url, headers={"User-Agent": "AbsoluteQuery-launcher"})
    with urllib.request.urlopen(req, timeout=timeout, context=TLS) as r:
        if progress is None:
            return r.read()
        total, parts, done = int(r.headers.get("Content-Length") or 0), [], 0
        while chunk := r.read(256 * 1024):
            parts.append(chunk)
            done += len(chunk)
            progress(done, total)
        return b"".join(parts)


class Progress:
    """Start-up window while an update downloads. Opens only when the update work is
    still running after SHOW_AFTER seconds, so a normal start shows nothing. The work
    runs in a thread and reports here; Tk is only touched from the main thread."""
    SHOW_AFTER = 0.4

    def __init__(self):
        self.text, self.done, self.total = "Checking for updates", 0, 0

    def stage(self, text: str) -> None:
        self.text, self.done, self.total = text, 0, 0

    def bytes(self, done: int, total: int) -> None:
        self.done, self.total = done, total

    def run(self, work: threading.Thread) -> None:
        work.join(self.SHOW_AFTER)
        if not work.is_alive():
            return
        try:
            self._window(work)
        except Exception:  # noqa: BLE001 - no display / no Tk: just wait
            work.join()

    def _window(self, work: threading.Thread) -> None:
        import tkinter as tk
        dark = _dark_theme()
        bg, fg, muted, track, bar = (("#1e1e1e", "#e6e6e6", "#9a9a9a", "#333333", "#4aa37a") if dark else
                                     ("#ffffff", "#1f1f1f", "#6b6b6b", "#e6e6e6", "#1e7a4f"))
        root = tk.Tk()
        root.title("AbsoluteQuery")
        root.configure(bg=bg)
        root.resizable(False, False)
        w, h = 380, 104
        root.geometry(f"{w}x{h}+{(root.winfo_screenwidth() - w) // 2}+{(root.winfo_screenheight() - h) // 3}")
        font = ("Segoe UI", 10) if os.name == "nt" else ("Sans", 10)
        stage = tk.Label(root, bg=bg, fg=fg, font=font, anchor="w")  # the window title names the app
        stage.place(x=24, y=24, width=w - 48)
        trough = tk.Frame(root, bg=track)  # plain frames: ttk bars ignore colours on some themes
        trough.place(x=24, y=52, width=w - 48, height=4)
        fill = tk.Frame(trough, bg=bar)
        detail = tk.Label(root, bg=bg, fg=muted, font=(font[0], 9), anchor="w")
        detail.place(x=24, y=68, width=w - 48)
        root.protocol("WM_DELETE_WINDOW", lambda: None)  # the update finishes on its own
        step = [0]

        def tick():
            if not work.is_alive():
                root.destroy()
                return
            stage.configure(text=self.text)
            if self.total:  # size known: how much of it has arrived
                fill.place(x=0, y=0, relheight=1, relwidth=min(1.0, self.done / self.total))
                detail.configure(text=f"{self.done / 1e6:.1f} of {self.total / 1e6:.1f} MB")
            else:  # size unknown: a short segment sweeps across
                step[0] = (step[0] + 1) % 40
                fill.place(relx=step[0] / 40 * 1.25 - 0.25, y=0, relheight=1, relwidth=0.25)
                detail.configure(text=f"{self.done / 1e6:.1f} MB" if self.done else "")
            root.after(40, tick)

        tick()
        root.mainloop()


def _dark_theme() -> bool:
    """The app's own theme setting, so this window matches what opens next."""
    try:
        return json.loads((Path.home() / ".absolutequery.json").read_text()).get("theme") == "Dark"
    except (OSError, ValueError, AttributeError):
        return False


def latest_sha(repo: str, branch: str) -> str:
    """Commit id at the tip of `branch`. Read from git's own ref listing rather than
    the GitHub API, which allows only 60 anonymous calls an hour per IP address -
    an office behind one IP would run out."""
    refs = get(f"{repo.rstrip('/')}.git/info/refs?service=git-upload-pack").decode("latin-1")
    m = re.search(rf"([0-9a-f]{{40}}) refs/heads/{re.escape(branch)}(?:\x00|\n|$)", refs)
    if not m:
        raise RuntimeError(f"branch {branch} not found")
    return m.group(1)


def download(repo: str, sha: str, ui: Progress) -> Path:
    """Unpack the repo at `sha` into versions/<sha>; returns the folder holding ENTRY.py."""
    dest = VERSIONS / sha[:12]
    if (dest / f"{ENTRY}.py").is_file():
        return dest
    ui.stage("Downloading the latest version")
    blob = get(f"{repo.rstrip('/')}/archive/{sha}.zip", DOWNLOAD_TIMEOUT, ui.bytes)
    ui.stage("Installing the latest version")
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


def update_program(state: dict, ui: Progress) -> bool:
    """Replace this program with the release's when that was built differently.
    True: replaced, restart it. Each new build is tried once, so a bad release cannot loop."""
    mine = build_id()
    if not mine:
        return False
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
            return False
        state["program_tried"] = latest
        write_state(state)
        ui.stage("Downloading a new program version")
        blob = get(f"{base}/{ASSET}", DOWNLOAD_TIMEOUT, ui.bytes)
        ui.stage("Installing, then restarting")
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
        return False
    return True


def restart() -> None:
    """Fresh start of the new program (not inside this one's unpacked files)."""
    exe = Path(sys.executable)
    subprocess.Popen([str(exe), *sys.argv[1:]], env={**os.environ, "PYINSTALLER_RESET_ENVIRONMENT": "1"})
    sys.exit(0)


def check_code():
    """Ask for the latest commit in the background while the program check runs.
    Returns a function that waits for the answer (or raises its error). A daemon
    thread, so a program restart does not wait for it."""
    repo = os.environ.get("ABSOLUTEQUERY_REPO", REPO_URL)
    branch = os.environ.get("ABSOLUTEQUERY_BRANCH", BRANCH)
    out: dict = {}

    def run():
        try:
            out["sha"] = latest_sha(repo, branch)
        except Exception as e:  # noqa: BLE001 - reported by update()
            out["error"] = e

    t = threading.Thread(target=run, daemon=True)
    t.start()

    def result() -> str:
        t.join()
        if "error" in out:
            raise out["error"]
        return out["sha"]

    return result


def update(state: dict, latest, ui: Progress) -> None:
    repo = os.environ.get("ABSOLUTEQUERY_REPO", REPO_URL)
    try:
        sha = latest()
        if sha == state.get("sha") and Path(state.get("dir", "")).is_dir():
            return
        d = download(repo, sha, ui)
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
    latest = check_code()  # both checks at once: one network wait, not two
    ui, out = Progress(), {}

    def work():
        out["restart"] = update_program(state, ui)
        if not out["restart"]:
            update(state, latest, ui)

    worker = threading.Thread(target=work, daemon=True)
    worker.start()
    ui.run(worker)  # a window only if this takes a while
    worker.join()
    if out.get("restart"):
        restart()
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
