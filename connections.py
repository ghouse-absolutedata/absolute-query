"""Saved server connections for AbsoluteQuery.

Server and user live in the settings file; passwords live only in the OS
keychain (Windows Credential Manager, macOS Keychain, Linux Secret Service)
through the `keyring` package. Without keyring, passwords are never stored.
"""
from __future__ import annotations

try:
    import keyring
    from keyring.errors import KeyringError
except ImportError:  # optional dependency
    keyring = None
    KeyringError = Exception

SERVICE = "AbsoluteQuery"


def keychain_ok() -> bool:
    if keyring is None:
        return False
    try:
        return keyring.get_keyring().priority > 0
    except Exception:  # noqa: BLE001 - any backend failure means unusable
        return False


def all_conns(settings: dict) -> list[dict]:
    return settings.setdefault("connections", [])


def find(settings: dict, name: str) -> dict | None:
    return next((c for c in all_conns(settings) if c["name"] == name), None)


def get_password(name: str) -> str:
    if not keychain_ok():
        return ""
    try:
        return keyring.get_password(SERVICE, name) or ""
    except KeyringError:
        return ""


def save(settings: dict, name: str, server: str, user: str, pwd: str, remember: bool,
         old_name: str | None = None) -> str | None:
    """Add or update a connection. Returns a warning when the password could not be kept."""
    if old_name and old_name != name:
        delete(settings, old_name)  # replaces the list, so read it afterwards
    conns = all_conns(settings)
    entry = find(settings, name)
    if entry is None:
        entry = {"name": name}
        conns.append(entry)
    entry.update(server=server, user=user, remember=bool(remember and pwd))
    conns.sort(key=lambda c: c["name"].lower())
    if not keychain_ok():
        entry["remember"] = False
        return "Password not saved: OS keychain unavailable (pip install keyring)" if remember else None
    try:
        if remember and pwd:
            keyring.set_password(SERVICE, name, pwd)
        else:
            _forget(name)
    except KeyringError as e:
        entry["remember"] = False
        return f"Password not saved: {e}"
    return None


def delete(settings: dict, name: str) -> None:
    settings["connections"] = [c for c in all_conns(settings) if c["name"] != name]
    if settings.get("last_conn") == name:
        settings.pop("last_conn", None)
    _forget(name)


def _forget(name: str) -> None:
    if keychain_ok():
        try:
            keyring.delete_password(SERVICE, name)
        except KeyringError:
            pass
