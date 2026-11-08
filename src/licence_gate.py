"""
Offline per-computer licence gate (the same model as the Pharmacy ERP).

The backend is the real gate. A licence is a plan: 30 days ($100) or lifetime ($1000).
  * Valid, or up to GRACE_DAYS after it ends: everything works (the UI shows a renewal notice).
  * Past the grace period: READ-ONLY. The person can still sign in, look at their data and
    download their reports and backups (it is their data), but every other call is refused
    (HTTP 402) with a message and the page that renews the licence.
  * No licence at all (never activated), a wrong-computer licence or a wound-back clock: locked,
    only the three licence methods and the page itself stay reachable.

Verification is Ed25519 against OWNER_PUBLIC_KEY (licence_owner_key.py, built into
every copy). A copy can check a licence but never make one. A build without a
valid public key refuses to run at all (fail closed); CI refuses to ship one
(src/licensing/check_owner_key.py).

State lives in a small JSON file in the data folder (licence-state.json). It is
separate from the database, never part of a backup or a restore, and never touches
business data.
"""

import os
import threading
import time

from licence_owner_key import OWNER_PUBLIC_KEY
from licensing.licensing import FileStore, Licensing

APP_ID = "DericBI Inventory Control"
PREFIX = "DIC"
STATE_FILENAME = "licence-state.json"
# Verifying a signature is cheap but not free; a minute of caching keeps the
# per-request cost at a dictionary lookup. Every activation resets it.
STATE_TTL_SECONDS = 60.0
GRACE_DAYS = 5
# Where a person renews: the DericBI site, pre-filled with this app and this computer's activation code.
RENEW_PAGE = "https://dericbi.vercel.app/renew.html"
RENEW_APP = "inventory-control"
# The only API methods that work while the copy is locked.
LICENCE_METHODS = frozenset({"licence_state", "activate_licence", "claim_owner"})
# What still works once the licence has run out: signing in (and the PIN recovery that goes with it),
# looking at data, and taking it out (reports, backups, Excel export). Nothing here changes business data.
READ_METHODS = frozenset({"is_pin_set", "verify_pin", "verify_recovery_key", "set_pin", "get_recovery_hint",
                          "lock_session", "log_action", "get_logo_path", "load_data", "get_config", "get_transfer_log",
                          "get_backups", "generate_pdf", "export_full_backup", "open_excel"})

_lock = threading.Lock()
_data_dir = None
_licensing = None
_cached = None


def set_data_dir(path):
    """Tell the gate which folder holds this copy's data (called once by main.py)."""
    global _data_dir
    with _lock:
        _data_dir = path


def get_licensing():
    global _licensing
    with _lock:
        if _licensing is None:
            folder = _data_dir or os.getcwd()
            _licensing = Licensing(APP_ID, PREFIX, OWNER_PUBLIC_KEY, FileStore(folder, STATE_FILENAME), grace_days=GRACE_DAYS)
        return _licensing


def use_licensing(licensing):
    """Swap the Licensing instance (tests use a throwaway key); None restores the real one."""
    global _licensing
    with _lock:
        _licensing = licensing
    reset_cache()


def reset_cache():
    global _cached
    with _lock:
        _cached = None


def public_state():
    """Licence state for the UI, plus the activation code needed to request or renew a licence."""
    lic = get_licensing()
    state = dict(lic.state())
    if lic.configured():
        state["install_id"] = lic.install_id()
        state["renew_url"] = renew_url(state["install_id"])
    return state


def renew_url(install_id):
    return f"{RENEW_PAGE}?app={RENEW_APP}&code={install_id}"


def current_state():
    """The cached answer to "may this copy be used right now?"."""
    global _cached
    now = time.monotonic()
    with _lock:
        cached = _cached
    if cached is not None and now - cached[0] < STATE_TTL_SECONDS:
        return cached[1]
    state = dict(get_licensing().state())
    with _lock:
        _cached = (now, state)
    return state


def allows(method, state):
    """May this call run? Everything for a valid licence; reading and exporting once it has lapsed; nothing otherwise."""
    if state["ok"]:
        return True
    return bool(state.get("read_only")) and (method is None or method in READ_METHODS)


def refusal(state, method=None):
    """The body sent with HTTP 402 when a call is refused for want of a licence."""
    out = {"ok": False, "licence": {"ok": False, "kind": state["kind"]}, "error": state["message"]}
    if state.get("read_only"):
        out["licence"]["read_only"] = True
        out["licence"]["renew_url"] = renew_url(state["install_id"])
        out["error"] = "Your licence has expired, so this is view-only. Renew your licence to add or change anything."
    return out
