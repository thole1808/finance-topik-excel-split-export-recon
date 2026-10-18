"""Local append-only, hash-chained usage history for the desktop app."""

from __future__ import annotations

import csv
import datetime as dt
import getpass
import hashlib
import json
import os
import platform
import socket
import sqlite3
from pathlib import Path
from typing import Any

from src.version import __version__
from src.device_identity import device_identity, masked_device_identity
from src.localization import format_local_datetime

ZERO_HASH = "0" * 64


def audit_directory() -> Path:
    if platform.system() == "Windows":
        root = Path(os.environ.get("APPDATA", Path.home() / "AppData/Roaming"))
        return root / "Recon Builder"
    if platform.system() == "Darwin":
        return Path.home() / "Library" / "Application Support" / "Recon Builder"
    return Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local/share")) / "recon-builder"


def generated_files_directory() -> Path:
    """Return the persistent per-user location for generated workbooks."""
    if platform.system() == "Windows":
        root = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData/Local"))
        return root / "Recon Builder" / "Generated Files"
    if platform.system() == "Darwin":
        return audit_directory() / "Generated Files"
    return audit_directory() / "Generated Files"


class AuditLog:
    """Persist app activity locally and detect edits to prior log entries."""

    def __init__(self, path: Path | None = None) -> None:
        self.path = path or audit_directory() / "usage-history.sqlite3"
        self.device_ip, self.device_mac = device_identity()
        self.device_ip_masked, self.device_mac_masked = masked_device_identity()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as connection:
            connection.execute("PRAGMA journal_mode=DELETE")
            connection.execute("PRAGMA synchronous=FULL")
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    occurred_at TEXT NOT NULL,
                    operator TEXT NOT NULL,
                    os_user TEXT NOT NULL,
                    host TEXT NOT NULL,
                    app_version TEXT NOT NULL,
                    action TEXT NOT NULL,
                    details_json TEXT NOT NULL,
                    previous_hash TEXT NOT NULL,
                    event_hash TEXT NOT NULL
                )
                """
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS events_time ON events(occurred_at DESC)"
            )

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path, timeout=15)
        connection.row_factory = sqlite3.Row
        return connection

    def last_operator(self) -> str:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT operator FROM events ORDER BY id DESC LIMIT 1"
            ).fetchone()
        return str(row[0]) if row else ""

    def last_successful_output(self) -> Path | None:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT details_json FROM events "
                "WHERE action = 'generate_succeeded' ORDER BY id DESC LIMIT 20"
            ).fetchall()
        for row in rows:
            output_path = json.loads(row[0]).get("output_path")
            if output_path:
                return Path(output_path)
        return None

    def generated_outputs(self, limit: int = 1000) -> list[dict[str, Any]]:
        """Return successful output records, newest first, for the results tab."""
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT occurred_at, details_json FROM events "
                "WHERE action = 'generate_succeeded' ORDER BY id DESC LIMIT ?",
                (limit,),
            ).fetchall()
        outputs: list[dict[str, Any]] = []
        for row in rows:
            details = json.loads(row["details_json"])
            output_path = details.get("output_path")
            if not output_path:
                continue
            outputs.append({
                "occurred_at": row["occurred_at"],
                "path": Path(output_path),
                "size_bytes": details.get("size_bytes", 0),
            })
        return outputs

    @staticmethod
    def _canonical(payload: dict[str, Any]) -> str:
        return json.dumps(
            payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
        )

    def record(self, operator: str, action: str, **details: Any) -> int:
        occurred_at = dt.datetime.now().astimezone().isoformat(timespec="seconds")
        details["ip_masked"] = self.device_ip_masked
        details["mac_masked"] = self.device_mac_masked
        details["ip_address"] = self.device_ip
        details["mac_address"] = self.device_mac
        details_json = self._canonical(details)
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            previous = connection.execute(
                "SELECT event_hash FROM events ORDER BY id DESC LIMIT 1"
            ).fetchone()
            previous_hash = str(previous[0]) if previous else ZERO_HASH
            payload = {
                "occurred_at": occurred_at,
                "operator": operator,
                "os_user": getpass.getuser(),
                "host": socket.gethostname(),
                "app_version": __version__,
                "action": action,
                "details_json": details_json,
                "previous_hash": previous_hash,
            }
            event_hash = hashlib.sha256(
                previous_hash.encode("ascii") + self._canonical(payload).encode("utf-8")
            ).hexdigest()
            cursor = connection.execute(
                """
                INSERT INTO events (
                    occurred_at, operator, os_user, host, app_version,
                    action, details_json, previous_hash, event_hash
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    occurred_at, operator, getpass.getuser(), socket.gethostname(),
                    __version__, action, details_json, previous_hash, event_hash,
                ),
            )
            return int(cursor.lastrowid)

    def verify_chain(self) -> tuple[bool, int | None]:
        previous_hash = ZERO_HASH
        with self._connect() as connection:
            rows = connection.execute("SELECT * FROM events ORDER BY id").fetchall()
        for row in rows:
            payload = {
                "occurred_at": row["occurred_at"],
                "operator": row["operator"],
                "os_user": row["os_user"],
                "host": row["host"],
                "app_version": row["app_version"],
                "action": row["action"],
                "details_json": row["details_json"],
                "previous_hash": previous_hash,
            }
            expected_hash = hashlib.sha256(
                previous_hash.encode("ascii") + self._canonical(payload).encode("utf-8")
            ).hexdigest()
            if row["previous_hash"] != previous_hash or row["event_hash"] != expected_hash:
                return False, int(row["id"])
            previous_hash = row["event_hash"]
        return True, None

    def recent_events(self, limit: int = 500) -> list[sqlite3.Row]:
        with self._connect() as connection:
            return list(connection.execute(
                "SELECT * FROM events ORDER BY id DESC LIMIT ?", (limit,)
            ).fetchall())

    def has_action(self, action: str) -> bool:
        with self._connect() as connection:
            return connection.execute(
                "SELECT 1 FROM events WHERE action = ? LIMIT 1", (action,)
            ).fetchone() is not None

    def export_csv(self, path: Path) -> Path:
        fields = (
            "id", "occurred_at", "operator", "os_user", "host", "app_version",
            "action", "ip_address", "mac_address", "details_json",
            "previous_hash", "event_hash",
        )
        with self._connect() as connection, Path(path).open(
            "w", encoding="utf-8-sig", newline=""
        ) as output:
            writer = csv.writer(output)
            writer.writerow(fields)
            for row in connection.execute(
                "SELECT " + ",".join(fields[:7] + ("details_json",) + fields[9:])
                + " FROM events ORDER BY id"
            ):
                details = json.loads(row["details_json"])
                row_values = {
                    **dict(row),
                    "ip_address": details.get("ip_address", ""),
                    "mac_address": details.get("mac_address", ""),
                    "details_json": self._canonical({
                        key: value for key, value in details.items()
                        if key not in {"ip_masked", "mac_masked"}
                    }),
                }
                writer.writerow([
                    format_local_datetime(row_values[field])
                    if field == "occurred_at" else row_values[field]
                    for field in fields
                ])
        return Path(path)
