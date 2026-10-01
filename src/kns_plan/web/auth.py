"""Shared office password with a lockout after repeated failures.

The password is never stored in plain text: the server keeps a salted PBKDF2 hash
(KNS_PASSWORD_HASH), made with `kns-plan hash-password`.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import os
from datetime import datetime, timedelta, timezone

from .store import Store

MAX_FAILURES = 5
WINDOW = timedelta(minutes=15)
ITERATIONS = 240_000


def hash_password(password: str, salt: bytes | None = None) -> str:
    salt = salt or os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, ITERATIONS)
    return f"pbkdf2${ITERATIONS}${base64.b64encode(salt).decode()}${base64.b64encode(digest).decode()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        _, iterations, salt, digest = stored.split("$")
        check = hashlib.pbkdf2_hmac("sha256", password.encode(), base64.b64decode(salt), int(iterations))
        return hmac.compare_digest(check, base64.b64decode(digest))
    except (ValueError, TypeError):
        return False


class Gate:
    def __init__(self, store: Store, password_hash: str):
        self.store = store
        self.password_hash = password_hash

    def locked_until(self, ip: str) -> datetime | None:
        since = (datetime.now(timezone.utc) - WINDOW).isoformat(timespec="seconds")
        failures = self.store.failures_since(ip, since)
        if len(failures) >= MAX_FAILURES:
            return datetime.fromisoformat(failures[-1]) + WINDOW
        return None

    def attempt(self, ip: str, password: str) -> tuple[bool, str]:
        until = self.locked_until(ip)
        if until:
            minutes = max(1, int((until - datetime.now(timezone.utc)).total_seconds() // 60) + 1)
            return False, f"Too many wrong passwords. Try again in {minutes} minute(s)."
        ok = verify_password(password, self.password_hash)
        self.store.record_login(ip, ok)
        if ok:
            return True, ""
        left = MAX_FAILURES - len(self.store.failures_since(
            ip, (datetime.now(timezone.utc) - WINDOW).isoformat(timespec="seconds")))
        return False, "Wrong password." + (f" {left} attempt(s) left." if left > 0 else " Locked for 15 minutes.")
