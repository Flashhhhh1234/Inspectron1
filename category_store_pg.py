"""PostgreSQL-backed category catalog storage helpers.

Loads and saves category definitions using the normalized category tables:
- category_types
- wiring_types
- subcategory_templates
- template_inputs
"""

from __future__ import annotations

import json
import os
import re
from collections import defaultdict
from typing import Any, Dict, Iterable, List, Sequence

import pg_sqlite_compat as sqlite3
from category_catalog_format import serialize_catalog

_DEFAULT_SCHEMA = "public"


def _config_paths() -> Sequence[str]:
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


def _owner_inputs(rows: Iterable[Any]) -> Dict[str, List[Dict[str, Any]]]:
    by_owner: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for row in rows:
        owner_uid = row["owner_uid"]
        by_owner[owner_uid].append(
            {
                "name": row["name"],
                "label": row["label"],
            }
        )
    return by_owner


def load_categories_from_postgres(db_key: str = "inspection_tool") -> List[Dict[str, Any]]:
    """Load categories from PostgreSQL normalized category tables."""
    schema = _load_schema()
    conn = sqlite3.connect(db_key)
    conn.row_factory = sqlite3.Row

    try:
        cur = conn.cursor()
        cur.execute(
            f"""
            SELECT category_uid, name, mode, ref_number, template, sort_order
            FROM {_qualified(schema, 'category_types')}
            ORDER BY sort_order, name
            """
        )
        cat_rows = cur.fetchall()

        cur.execute(
            f"""
            SELECT wiring_uid, category_uid, type, ref_number, sort_order
            FROM {_qualified(schema, 'wiring_types')}
            ORDER BY category_uid, sort_order, type
            """
        )
        wiring_rows = cur.fetchall()

        cur.execute(
            f"""
            SELECT subcategory_uid, category_uid, wiring_uid, scope, name, ref_number, template, sort_order
            FROM {_qualified(schema, 'subcategory_templates')}
            ORDER BY category_uid, COALESCE(wiring_uid, ''), sort_order, name
            """
        )
        sub_rows = cur.fetchall()

        cur.execute(
            f"""
            SELECT input_uid, owner_scope, owner_uid, name, label, sort_order
            FROM {_qualified(schema, 'template_inputs')}
            ORDER BY owner_uid, sort_order, name
            """
        )
        input_rows = cur.fetchall()
    except Exception as exc:
        conn.close()
        print(f"Error loading categories from PostgreSQL: {exc}")
        return []

    conn.close()

    inputs_by_owner = _owner_inputs(input_rows)
    wiring_by_category: Dict[str, List[Any]] = defaultdict(list)
    for row in wiring_rows:
        wiring_by_category[row["category_uid"]].append(row)

    subs_by_category: Dict[str, List[Any]] = defaultdict(list)
    subs_by_wiring: Dict[str, List[Any]] = defaultdict(list)
    special_by_category: Dict[str, List[Any]] = defaultdict(list)

    for row in sub_rows:
        scope = row["scope"]
        if scope == "wiring" and row["wiring_uid"]:
            subs_by_wiring[row["wiring_uid"]].append(row)
        elif scope == "special":
            special_by_category[row["category_uid"]].append(row)
        else:
            subs_by_category[row["category_uid"]].append(row)

    def to_sub_dict(row: Any) -> Dict[str, Any]:
        sub_uid = row["subcategory_uid"]
        return {
            "name": row["name"],
            "ref_number": row["ref_number"],
            "inputs": inputs_by_owner.get(sub_uid, []),
            "template": row["template"],
        }

    categories: List[Dict[str, Any]] = []

    for row in cat_rows:
        cat_uid = row["category_uid"]
        mode = row["mode"] or ("wiring_selector" if wiring_by_category.get(cat_uid) else "parent")

        category: Dict[str, Any] = {
            "name": row["name"],
            "mode": mode,
            "ref_number": row["ref_number"],
            "inputs": inputs_by_owner.get(cat_uid, []),
            "template": row["template"],
        }

        if mode == "wiring_selector":
            wiring_items = []
            for wiring in wiring_by_category.get(cat_uid, []):
                wiring_uid = wiring["wiring_uid"]
                wiring_items.append(
                    {
                        "type": wiring["type"],
                        "ref_number": wiring["ref_number"],
                        "subcategories": [to_sub_dict(sub) for sub in subs_by_wiring.get(wiring_uid, [])],
                    }
                )

            category["wiring_types"] = wiring_items
            category["special_subcategories"] = [
                to_sub_dict(sub) for sub in special_by_category.get(cat_uid, [])
            ]
        else:
            category["subcategories"] = [to_sub_dict(sub) for sub in subs_by_category.get(cat_uid, [])]

        categories.append(category)

    return categories


def _upsert_rows(
    cursor: Any,
    schema: str,
    table: str,
    columns: Sequence[str],
    pk_column: str,
    rows: List[Dict[str, Any]],
) -> None:
    if not rows:
        return

    col_sql = ", ".join(_q(col) for col in columns)
    placeholders = ", ".join(["?"] * len(columns))
    updates = [col for col in columns if col != pk_column]
    update_sql = ", ".join(f"{_q(col)} = EXCLUDED.{_q(col)}" for col in updates)

    sql = (
        f"INSERT INTO {_qualified(schema, table)} ({col_sql}) VALUES ({placeholders}) "
        f"ON CONFLICT ({_q(pk_column)}) DO UPDATE SET {update_sql}"
    )
    values = [tuple(row.get(col) for col in columns) for row in rows]
    cursor.executemany(sql, values)


def _delete_missing(
    cursor: Any,
    schema: str,
    table: str,
    pk_column: str,
    keep_ids: List[Any],
) -> None:
    if not keep_ids:
        cursor.execute(f"DELETE FROM {_qualified(schema, table)}")
        return

    placeholders = ", ".join(["?"] * len(keep_ids))
    cursor.execute(
        f"DELETE FROM {_qualified(schema, table)} WHERE {_q(pk_column)} NOT IN ({placeholders})",
        tuple(keep_ids),
    )


def save_categories_to_postgres(categories: List[Dict[str, Any]], db_key: str = "inspection_tool") -> None:
    """Persist categories into PostgreSQL normalized category tables."""
    payload = serialize_catalog(categories, source="postgres_save")
    seed = payload.get("postgres_seed", {}) if isinstance(payload, dict) else {}

    category_types = list(seed.get("category_types", []))
    wiring_types = list(seed.get("wiring_types", []))
    subcategory_templates = list(seed.get("subcategory_templates", []))
    template_inputs = list(seed.get("template_inputs", []))

    schema = _load_schema()

    conn = sqlite3.connect(db_key)
    try:
        cur = conn.cursor()

        _upsert_rows(
            cur,
            schema,
            "category_types",
            ["category_uid", "name", "mode", "ref_number", "template", "sort_order"],
            "category_uid",
            category_types,
        )
        _upsert_rows(
            cur,
            schema,
            "wiring_types",
            ["wiring_uid", "category_uid", "type", "ref_number", "sort_order"],
            "wiring_uid",
            wiring_types,
        )
        _upsert_rows(
            cur,
            schema,
            "subcategory_templates",
            [
                "subcategory_uid",
                "category_uid",
                "wiring_uid",
                "scope",
                "name",
                "ref_number",
                "template",
                "sort_order",
            ],
            "subcategory_uid",
            subcategory_templates,
        )
        _upsert_rows(
            cur,
            schema,
            "template_inputs",
            ["input_uid", "owner_scope", "owner_uid", "name", "label", "sort_order"],
            "input_uid",
            template_inputs,
        )

        _delete_missing(
            cur,
            schema,
            "template_inputs",
            "input_uid",
            [row.get("input_uid") for row in template_inputs],
        )
        _delete_missing(
            cur,
            schema,
            "subcategory_templates",
            "subcategory_uid",
            [row.get("subcategory_uid") for row in subcategory_templates],
        )
        _delete_missing(
            cur,
            schema,
            "wiring_types",
            "wiring_uid",
            [row.get("wiring_uid") for row in wiring_types],
        )
        _delete_missing(
            cur,
            schema,
            "category_types",
            "category_uid",
            [row.get("category_uid") for row in category_types],
        )

        conn.commit()
    except Exception as exc:
        conn.rollback()
        raise RuntimeError(
            "Failed to save categories to PostgreSQL. Ensure category tables are imported first."
        ) from exc
    finally:
        conn.close()
