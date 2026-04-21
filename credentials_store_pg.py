"""PostgreSQL-backed credential storage helpers."""

from __future__ import annotations

import json
import os
import re
from typing import Any, Dict

import pg_sqlite_compat as sqlite3

_DEFAULT_SCHEMA = "public"
_DEFAULT_DB_KEY = "inspection_tool"
_TABLE_NAME = "user_credentials"


def _config_paths():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    return (
        os.path.join(current_dir, "assets", "postgres.json"),
        os.path.join(os.path.dirname(current_dir), "assets", "postgres.json"),
    )


def _safe_schema_name(candidate: Any) -> str:
    value = str(candidate or "").strip()
    if re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", value):
        return value
    return _DEFAULT_SCHEMA


def _load_schema() -> str:
    for config_path in _config_paths():
        if not os.path.exists(config_path):
            continue

        try:
            with open(config_path, "r", encoding="utf-8") as handle:
                raw = json.load(handle)
        except Exception:
            continue

        if not isinstance(raw, dict):
            continue

        block = raw.get("postgres", raw)
        if isinstance(block, dict):
            schema = block.get("schema") or raw.get("schema")
            return _safe_schema_name(schema)

    return _DEFAULT_SCHEMA


def _q(name: str) -> str:
    return '"' + str(name).replace('"', '""') + '"'


def _qualified(schema: str, table: str) -> str:
    return f"{_q(schema)}.{_q(table)}"


def load_users_from_postgres(db_key: str = _DEFAULT_DB_KEY) -> Dict[str, Dict[str, str]]:
    """Load active users from PostgreSQL credential table."""
    schema = _load_schema()
    conn = sqlite3.connect(db_key)
    conn.row_factory = sqlite3.Row

    try:
        cur = conn.cursor()
        cur.execute(
            f"""
            SELECT username, password, role, full_name
            FROM {_qualified(schema, _TABLE_NAME)}
            WHERE COALESCE(is_active, TRUE) = TRUE
            ORDER BY username
            """
        )
        rows = cur.fetchall()
    except Exception as exc:
        raise RuntimeError(
            "Failed to load credentials from PostgreSQL. "
            "Run scripts/import_credentials_to_postgres.py first."
        ) from exc
    finally:
        conn.close()

    users: Dict[str, Dict[str, str]] = {}
    for row in rows:
        username = row["username"]
        users[username] = {
            "password": row["password"],
            "role": row["role"],
            "full_name": row["full_name"] or username,
        }

    return users


def save_users_to_postgres(users: Dict[str, Dict[str, str]], db_key: str = _DEFAULT_DB_KEY) -> None:
    """Replace PostgreSQL credential rows with provided user mapping."""
    schema = _load_schema()
    conn = sqlite3.connect(db_key)

    upsert_sql = (
        f"INSERT INTO {_qualified(schema, _TABLE_NAME)} "
        f"({', '.join([_q('username'), _q('password'), _q('role'), _q('full_name'), _q('is_active')])}) "
        f"VALUES (?, ?, ?, ?, ?) "
        f"ON CONFLICT ({_q('username')}) DO UPDATE SET "
        f"{_q('password')} = EXCLUDED.{_q('password')}, "
        f"{_q('role')} = EXCLUDED.{_q('role')}, "
        f"{_q('full_name')} = EXCLUDED.{_q('full_name')}, "
        f"{_q('is_active')} = EXCLUDED.{_q('is_active')}"
    )

    usernames = sorted(users.keys())

    try:
        cur = conn.cursor()

        rows = []
        for username in usernames:
            data = users.get(username) or {}
            rows.append(
                (
                    username,
                    str(data.get("password", "")),
                    str(data.get("role", "Quality")),
                    str(data.get("full_name") or username),
                    True,
                )
            )

        if rows:
            cur.executemany(upsert_sql, rows)

        if usernames:
            placeholders = ", ".join(["?"] * len(usernames))
            cur.execute(
                f"DELETE FROM {_qualified(schema, _TABLE_NAME)} WHERE {_q('username')} NOT IN ({placeholders})",
                tuple(usernames),
            )
        else:
            cur.execute(f"DELETE FROM {_qualified(schema, _TABLE_NAME)}")

        conn.commit()
    except Exception as exc:
        conn.rollback()
        raise RuntimeError(
            "Failed to save credentials into PostgreSQL. "
            "Run scripts/import_credentials_to_postgres.py first."
        ) from exc
    finally:
        conn.close()
