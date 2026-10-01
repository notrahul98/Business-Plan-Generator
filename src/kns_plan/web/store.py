"""Plan storage. Every save is a new version, so nothing is ever overwritten.

SQLite by default (one file, easy to back up). The interface is small on purpose so a SQL Server or
PostgreSQL version can replace it once IT confirms the server.
"""

from __future__ import annotations

import json
import shutil
import sqlite3
import threading
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from ..schema import Plan

SCHEMA = """
CREATE TABLE IF NOT EXISTS plans (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    brand TEXT NOT NULL,
    archived INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS versions (
    plan_id INTEGER NOT NULL REFERENCES plans(id),
    version INTEGER NOT NULL,
    data TEXT NOT NULL,
    saved_at TEXT NOT NULL,
    saved_by TEXT NOT NULL DEFAULT '',
    note TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (plan_id, version)
);
CREATE TABLE IF NOT EXISTS login_attempts (
    ip TEXT NOT NULL,
    at TEXT NOT NULL,
    ok INTEGER NOT NULL
);
"""


@dataclass
class PlanRow:
    id: int
    name: str
    brand: str
    version: int
    updated_at: str
    archived: bool


@dataclass
class VersionRow:
    version: int
    saved_at: str
    saved_by: str
    note: str


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class Store:
    def __init__(self, path: Path | str):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        with self._conn() as c:
            c.executescript(SCHEMA)

    @contextmanager
    def _conn(self):
        conn = sqlite3.connect(self.path, timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    # plans ---------------------------------------------------------
    def list_plans(self, include_archived: bool = False) -> list[PlanRow]:
        with self._conn() as c:
            rows = c.execute(
                "SELECT p.*, (SELECT MAX(version) FROM versions v WHERE v.plan_id = p.id) AS version "
                "FROM plans p WHERE archived = 0 OR ? ORDER BY updated_at DESC", (include_archived,)).fetchall()
        return [PlanRow(r["id"], r["name"], r["brand"], r["version"], r["updated_at"], bool(r["archived"]))
                for r in rows]

    def create(self, name: str, plan: Plan, by: str = "") -> int:
        with self._lock, self._conn() as c:
            t = now()
            cur = c.execute("INSERT INTO plans (name, brand, created_at, updated_at) VALUES (?, ?, ?, ?)",
                            (name, plan.settings.brand, t, t))
            pid = cur.lastrowid
            c.execute("INSERT INTO versions VALUES (?, 1, ?, ?, ?, ?)",
                      (pid, plan.model_dump_json(exclude_none=True), t, by, "created"))
        return pid

    def get(self, plan_id: int, version: int | None = None) -> tuple[PlanRow, Plan]:
        with self._conn() as c:
            p = c.execute("SELECT * FROM plans WHERE id = ?", (plan_id,)).fetchone()
            if p is None:
                raise KeyError(plan_id)
            v = c.execute("SELECT * FROM versions WHERE plan_id = ? AND (? IS NULL OR version = ?) "
                          "ORDER BY version DESC LIMIT 1", (plan_id, version, version)).fetchone()
            if v is None:
                raise KeyError((plan_id, version))
        row = PlanRow(p["id"], p["name"], p["brand"], v["version"], p["updated_at"], bool(p["archived"]))
        return row, Plan.model_validate_json(v["data"])

    def save(self, plan_id: int, plan: Plan, by: str = "", note: str = "") -> int:
        with self._lock, self._conn() as c:
            latest = c.execute("SELECT MAX(version) FROM versions WHERE plan_id = ?", (plan_id,)).fetchone()[0]
            if latest is None:
                raise KeyError(plan_id)
            data = plan.model_dump_json(exclude_none=True)
            prev = c.execute("SELECT data FROM versions WHERE plan_id = ? AND version = ?",
                             (plan_id, latest)).fetchone()[0]
            if data == prev:
                return latest  # nothing changed
            t = now()
            c.execute("INSERT INTO versions VALUES (?, ?, ?, ?, ?, ?)", (plan_id, latest + 1, data, t, by, note))
            c.execute("UPDATE plans SET updated_at = ?, brand = ? WHERE id = ?", (t, plan.settings.brand, plan_id))
        return latest + 1

    def rename(self, plan_id: int, name: str) -> None:
        with self._conn() as c:
            c.execute("UPDATE plans SET name = ? WHERE id = ?", (name, plan_id))

    def archive(self, plan_id: int, archived: bool = True) -> None:
        with self._conn() as c:
            c.execute("UPDATE plans SET archived = ? WHERE id = ?", (int(archived), plan_id))

    def versions(self, plan_id: int) -> list[VersionRow]:
        with self._conn() as c:
            rows = c.execute("SELECT version, saved_at, saved_by, note FROM versions WHERE plan_id = ? "
                             "ORDER BY version DESC", (plan_id,)).fetchall()
        return [VersionRow(r["version"], r["saved_at"], r["saved_by"], r["note"]) for r in rows]

    def export_json(self, plan_id: int) -> str:
        row, plan = self.get(plan_id)
        return json.dumps({"name": row.name, "version": row.version, "plan": plan.model_dump(exclude_none=True)},
                          indent=2)

    # logins --------------------------------------------------------
    def record_login(self, ip: str, ok: bool) -> None:
        with self._conn() as c:
            c.execute("INSERT INTO login_attempts VALUES (?, ?, ?)", (ip, now(), int(ok)))

    def failures_since(self, ip: str, since: str) -> list[str]:
        """Times of failed logins from this IP since `since`, after its last successful login."""
        with self._conn() as c:
            last_ok = c.execute("SELECT MAX(at) FROM login_attempts WHERE ip = ? AND ok = 1", (ip,)).fetchone()[0]
            rows = c.execute("SELECT at FROM login_attempts WHERE ip = ? AND ok = 0 AND at >= ? AND at > ? "
                             "ORDER BY at", (ip, since, last_ok or "")).fetchall()
        return [r[0] for r in rows]

    # backups -------------------------------------------------------
    def backup(self, folder: Path | str, keep: int = 30) -> Path:
        """Consistent copy of the database, e.g. from a nightly scheduled task. Keeps the newest `keep` copies."""
        folder = Path(folder)
        folder.mkdir(parents=True, exist_ok=True)
        target = folder / f"plans-{datetime.now():%Y%m%d-%H%M%S}.sqlite"
        with self._conn() as c:
            dest = sqlite3.connect(target)
            c.backup(dest)
            dest.close()
        for old in sorted(folder.glob("plans-*.sqlite"))[:-keep]:
            old.unlink()
        return target


def copy_file(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(src, dst)
