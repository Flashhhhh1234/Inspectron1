"""Compatibility layer to let sqlite3-style code run on PostgreSQL.

This module intentionally exposes a small sqlite3-like surface used by this
codebase: connect, Row, IntegrityError, OperationalError, and basic
Connection/Cursor behavior.
"""

from __future__ import annotations

import json
import os
import re
import importlib
import sys
from typing import Iterable, Optional, Sequence

try:
    _pg = importlib.import_module("psycopg2")
    _pg_driver = "psycopg2"
except ImportError:  # pragma: no cover
    try:
        _pg = importlib.import_module("psycopg")
        _pg_driver = "psycopg"
    except ImportError:  # pragma: no cover
        _pg = None
        _pg_driver = None


if _pg is not None:
    IntegrityError = _pg.IntegrityError
    OperationalError = _pg.Error
    DatabaseError = _pg.Error
else:  # pragma: no cover
    class DatabaseError(Exception):
        pass

    class IntegrityError(DatabaseError):
        pass

    class OperationalError(DatabaseError):
        pass


class Row(dict):
    """sqlite3.Row-like mapping that also supports integer indexing."""

    def __init__(self, pairs):
        super().__init__(pairs)
        self._ordered_values = list(super().values())

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._ordered_values[key]
        return super().__getitem__(key)


def _load_postgres_config() -> dict:
    """Load optional PostgreSQL settings from assets/postgres.json."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    config_paths = (
        os.path.join(current_dir, "assets", "postgres.json"),
        os.path.join(os.path.dirname(current_dir), "assets", "postgres.json"),
    )

    key_aliases = {
        "url": "database_url",
        "database": "dbname",
        "db": "dbname",
        "username": "user",
        "pass": "password",
    }

    merged = {}

    for config_path in config_paths:
        if not os.path.exists(config_path):
            continue

        try:
            with open(config_path, "r", encoding="utf-8") as file:
                raw = json.load(file)
        except Exception:
            continue

        block = raw.get("postgres", raw) if isinstance(raw, dict) else {}
        if not isinstance(block, dict):
            continue

        normalized = {}
        for raw_key, raw_value in block.items():
            if raw_value is None:
                continue
            key = key_aliases.get(str(raw_key).strip().lower(), str(raw_key).strip().lower())
            value = str(raw_value).strip()
            if value:
                normalized[key] = value

        for key, value in normalized.items():
            if key not in merged or merged[key] in (None, ""):
                merged[key] = value

    return merged


def _setting(config: dict, env_primary: str, env_secondary: Optional[str] = None, default: Optional[str] = None):
    for env_name in (env_primary, env_secondary):
        if not env_name:
            continue
        value = os.getenv(env_name)
        if value not in (None, ""):
            return value

    for key in (env_primary.lower(), (env_secondary or "").lower()):
        if key and key in config and config[key] not in (None, ""):
            return config[key]

    return default


def _schema_from_db_path(db_path: Optional[str]) -> str:
    if not db_path:
        return os.getenv("POSTGRES_SCHEMA", "public")

    base = os.path.splitext(os.path.basename(str(db_path)))[0].strip().lower()
    schema = re.sub(r"[^a-z0-9_]", "_", base)
    if not schema:
        schema = "public"
    if schema[0].isdigit():
        schema = f"s_{schema}"
    return schema


def _quote_identifier(identifier: str) -> str:
    return '"' + identifier.replace('"', '""') + '"'


def _is_missing_database_error(exc: Exception, dbname: str) -> bool:
    message = str(exc).lower()
    return "does not exist" in message and f'database "{dbname.lower()}"' in message


def _ensure_database_exists(host: str, port: int, dbname: str, user: str, password: Optional[str], maintenance_db: str):
    admin_kwargs = {
        "host": host,
        "port": port,
        "dbname": maintenance_db,
        "user": user,
    }
    if password not in (None, ""):
        admin_kwargs["password"] = password

    admin_conn = _pg.connect(**admin_kwargs)
    try:
        admin_conn.autocommit = True
        with admin_conn.cursor() as cur:
            cur.execute("SELECT 1 FROM pg_database WHERE datname = %s", (dbname,))
            if cur.fetchone() is None:
                cur.execute(f"CREATE DATABASE {_quote_identifier(dbname)}")
    finally:
        admin_conn.close()


def _connect_postgres(db_path: Optional[str]):
    if _pg is None:
        raise ModuleNotFoundError(
            "PostgreSQL driver not installed in this interpreter. "
            f"Install with: \"{sys.executable}\" -m pip install psycopg2-binary"
        )

    config = _load_postgres_config()
    host = _setting(config, "POSTGRES_HOST", "PGHOST", "localhost")
    port = _setting(config, "POSTGRES_PORT", "PGPORT", "5432")
    dbname = _setting(config, "POSTGRES_DB", "PGDATABASE", config.get("dbname", "inspection_tool"))
    user = _setting(config, "POSTGRES_USER", "PGUSER", "postgres")
    password = _setting(config, "POSTGRES_PASSWORD", "PGPASSWORD", config.get("password"))
    maintenance_db = _setting(config, "POSTGRES_MAINTENANCE_DB", default=config.get("maintenance_db", "postgres"))

    database_url = _setting(config, "DATABASE_URL", default=config.get("database_url"))
    if database_url:
        dsn_overrides = {
            "host": host,
            "port": int(port),
            "dbname": dbname,
            "user": user,
        }
        if password not in (None, ""):
            dsn_overrides["password"] = password

        try:
            conn = _pg.connect(database_url, **dsn_overrides)
        except Exception as exc:
            if not _is_missing_database_error(exc, dbname):
                raise

            _ensure_database_exists(
                host=host,
                port=int(port),
                dbname=dbname,
                user=user,
                password=password,
                maintenance_db=maintenance_db,
            )
            conn = _pg.connect(database_url, **dsn_overrides)
    else:
        connect_kwargs = {
            "host": host,
            "port": int(port),
            "dbname": dbname,
            "user": user,
        }
        if password not in (None, ""):
            connect_kwargs["password"] = password

        try:
            conn = _pg.connect(
                **connect_kwargs,
            )
        except Exception as exc:
            if not _is_missing_database_error(exc, dbname):
                raise

            _ensure_database_exists(
                host=host,
                port=int(port),
                dbname=dbname,
                user=user,
                password=password,
                maintenance_db=maintenance_db,
            )
            conn = _pg.connect(
                **connect_kwargs,
            )

    conn.autocommit = False

    schema = _setting(config, "POSTGRES_SCHEMA", default=config.get("schema")) or _schema_from_db_path(db_path)
    if schema and schema != "public":
        with conn.cursor() as cur:
            cur.execute(f'CREATE SCHEMA IF NOT EXISTS "{schema}"')
            cur.execute(f'SET search_path TO "{schema}", public')
        conn.commit()

    return conn


def _replace_qmark_placeholders(sql: str) -> str:
    """Replace sqlite qmark placeholders (?) with psycopg placeholders (%s)."""
    out = []
    in_single = False
    in_double = False

    for ch in sql:
        if ch == "'" and not in_double:
            in_single = not in_single
            out.append(ch)
            continue
        if ch == '"' and not in_single:
            in_double = not in_double
            out.append(ch)
            continue
        if ch == "?" and not in_single and not in_double:
            out.append("%s")
        else:
            out.append(ch)

    return "".join(out)


def _rewrite_insert_or_replace(sql: str) -> str:
    pattern = re.compile(
        r"^\s*INSERT\s+OR\s+REPLACE\s+INTO\s+([a-zA-Z_][\w]*)\s*"
        r"\((.*?)\)\s*VALUES\s*\((.*?)\)\s*$",
        re.IGNORECASE | re.DOTALL,
    )
    match = pattern.match(sql)
    if not match:
        return sql

    table = match.group(1)
    columns_str = match.group(2)
    values_str = match.group(3)

    columns = [c.strip() for c in columns_str.split(",") if c.strip()]
    conflict_col = None
    for col in columns:
        if col.lower() == "cabinet_id":
            conflict_col = col
            break

    if not conflict_col:
        return sql.replace("INSERT OR REPLACE", "INSERT", 1)

    updates = [f"{col} = EXCLUDED.{col}" for col in columns if col.lower() != conflict_col.lower()]
    update_clause = ", ".join(updates) if updates else f"{conflict_col} = EXCLUDED.{conflict_col}"

    return (
        f"INSERT INTO {table} ({columns_str}) VALUES ({values_str}) "
        f"ON CONFLICT ({conflict_col}) DO UPDATE SET {update_clause}"
    )


def _rewrite_insert_or_ignore(sql: str) -> str:
    pattern = re.compile(r"^\s*INSERT\s+OR\s+IGNORE\s+INTO\s+", re.IGNORECASE)
    if not pattern.match(sql):
        return sql

    stripped = sql.strip().rstrip(";")
    rewritten = pattern.sub("INSERT INTO ", stripped, count=1)
    return f"{rewritten} ON CONFLICT DO NOTHING"


def _rewrite_add_column_if_not_exists(sql: str) -> str:
    pattern = re.compile(
        r"^\s*ALTER\s+TABLE\s+([a-zA-Z_][\w]*)\s+ADD\s+COLUMN\s+(?!IF\s+NOT\s+EXISTS)(.+?)\s*$",
        re.IGNORECASE | re.DOTALL,
    )
    match = pattern.match(sql)
    if not match:
        return sql

    table = match.group(1)
    column_clause = match.group(2).strip()
    return f"ALTER TABLE {table} ADD COLUMN IF NOT EXISTS {column_clause}"


def _transform_sql(sql: str) -> str:
    transformed = sql

    transformed = re.sub(
        r"\bINTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT\b",
        "SERIAL PRIMARY KEY",
        transformed,
        flags=re.IGNORECASE,
    )
    transformed = re.sub(r"\bAUTOINCREMENT\b", "", transformed, flags=re.IGNORECASE)
    transformed = _rewrite_insert_or_replace(transformed)
    transformed = _rewrite_insert_or_ignore(transformed)
    transformed = _rewrite_add_column_if_not_exists(transformed)
    transformed = re.sub(
        r"\bDATE\s*\(\s*([a-zA-Z_][\w]*)\s*\)",
        r"CAST(\1 AS DATE)",
        transformed,
        flags=re.IGNORECASE,
    )
    transformed = _replace_qmark_placeholders(transformed)

    return transformed


class Cursor:
    def __init__(self, connection: "Connection", cursor):
        self._connection = connection
        self._cursor = cursor

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False

    def execute(self, sql: str, params: Optional[Sequence] = None):
        sql = _transform_sql(sql)
        if params is None:
            self._cursor.execute(sql)
        else:
            self._cursor.execute(sql, tuple(params))
        return self

    def executemany(self, sql: str, seq_of_params: Iterable[Sequence]):
        sql = _transform_sql(sql)
        normalized = [tuple(params) for params in seq_of_params]
        self._cursor.executemany(sql, normalized)
        return self

    def fetchone(self):
        row = self._cursor.fetchone()
        return self._convert_row(row)

    def fetchall(self):
        rows = self._cursor.fetchall()
        return [self._convert_row(row) for row in rows]

    @property
    def rowcount(self):
        return self._cursor.rowcount

    def close(self):
        self._cursor.close()

    def _convert_row(self, row):
        if row is None:
            return None

        if self._connection.row_factory is Row:
            columns = [desc[0] for desc in (self._cursor.description or [])]
            return Row(zip(columns, row))

        return row


class Connection:
    def __init__(self, raw_connection):
        self._connection = raw_connection
        self.row_factory = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        if exc_type is None:
            self.commit()
        else:
            self.rollback()
        self.close()
        return False

    def cursor(self):
        return Cursor(self, self._connection.cursor())

    def commit(self):
        self._connection.commit()

    def rollback(self):
        self._connection.rollback()

    def close(self):
        self._connection.close()


def connect(db_path: Optional[str] = None):
    """sqlite-style connect signature; db_path is used only to derive schema."""
    return Connection(_connect_postgres(db_path))
