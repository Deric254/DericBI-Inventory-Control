"""
Offline per-computer licence gate (the same model as the Pharmacy ERP).

The backend is the real gate: every API call and every report/backup download is
refused (HTTP 402) unless this copy holds a valid licence, or is the owner's own
computer. The UI only shows the licence screen. Only the three licence methods,
the page itself, its fonts and the logo stay reachable while locked; none of them
expose business data.

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
# The only API methods that work while the copy is locked.
LICENCE_METHODS = frozenset({"licence_state", "activate_licence", "claim_owner"})

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
            _licensing = Licensing(APP_ID, PREFIX, OWNER_PUBLIC_KEY, FileStore(folder, STATE_FILENAME))
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
    return state


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


def refusal(state):
    """The body sent with HTTP 402 when a call is refused for want of a licence."""
    return {"ok": False, "licence": {"ok": False, "kind": state["kind"]}, "error": state["message"]}
