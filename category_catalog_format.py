"""Category catalog JSON format helpers.

Supports:
- Backward-compatible load from legacy list format.
- Serialization to a PostgreSQL-ready envelope for future migration.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any, Dict, List, Tuple
import re


def _slug(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text or "item"


def extract_categories(payload: Any) -> Tuple[List[Dict[str, Any]], bool]:
    """Return (categories, already_enveloped)."""
    if isinstance(payload, list):
        return payload, False

    if isinstance(payload, dict):
        categories = payload.get("categories")
        if isinstance(categories, list):
            return categories, True

    return [], False


def _build_postgres_seed(categories: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    seed = {
        "category_types": [],
        "wiring_types": [],
        "subcategory_templates": [],
        "template_inputs": [],
    }

    for cat_idx, category in enumerate(categories, 1):
        cat_name = category.get("name", "")
        cat_uid = f"cat_{cat_idx}_{_slug(cat_name)}"

        seed["category_types"].append(
            {
                "category_uid": cat_uid,
                "name": cat_name,
                "mode": category.get("mode"),
                "ref_number": category.get("ref_number"),
                "template": category.get("template"),
                "sort_order": cat_idx,
            }
        )

        category_inputs = category.get("inputs", [])
        if isinstance(category_inputs, list):
            for inp_idx, inp in enumerate(category_inputs, 1):
                seed["template_inputs"].append(
                    {
                        "input_uid": f"inp_{cat_uid}_{inp_idx}",
                        "owner_scope": "category",
                        "owner_uid": cat_uid,
                        "name": inp.get("name"),
                        "label": inp.get("label"),
                        "sort_order": inp_idx,
                    }
                )

        wiring_types = category.get("wiring_types", [])
        if isinstance(wiring_types, list):
            for wire_idx, wiring in enumerate(wiring_types, 1):
                wiring_uid = f"wire_{cat_uid}_{wire_idx}_{_slug(wiring.get('type'))}"
                seed["wiring_types"].append(
                    {
                        "wiring_uid": wiring_uid,
                        "category_uid": cat_uid,
                        "type": wiring.get("type"),
                        "ref_number": wiring.get("ref_number"),
                        "sort_order": wire_idx,
                    }
                )

                for sub_idx, sub in enumerate(wiring.get("subcategories", []), 1):
                    sub_uid = f"sub_{wiring_uid}_{sub_idx}_{_slug(sub.get('name'))}"
                    seed["subcategory_templates"].append(
                        {
                            "subcategory_uid": sub_uid,
                            "category_uid": cat_uid,
                            "wiring_uid": wiring_uid,
                            "scope": "wiring",
                            "name": sub.get("name"),
                            "ref_number": sub.get("ref_number"),
                            "template": sub.get("template"),
                            "sort_order": sub_idx,
                        }
                    )

                    for inp_idx, inp in enumerate(sub.get("inputs", []), 1):
                        seed["template_inputs"].append(
                            {
                                "input_uid": f"inp_{sub_uid}_{inp_idx}",
                                "owner_scope": "subcategory",
                                "owner_uid": sub_uid,
                                "name": inp.get("name"),
                                "label": inp.get("label"),
                                "sort_order": inp_idx,
                            }
                        )

        regular_subs = category.get("subcategories", [])
        if isinstance(regular_subs, list):
            for sub_idx, sub in enumerate(regular_subs, 1):
                sub_uid = f"sub_{cat_uid}_{sub_idx}_{_slug(sub.get('name'))}"
                seed["subcategory_templates"].append(
                    {
                        "subcategory_uid": sub_uid,
                        "category_uid": cat_uid,
                        "wiring_uid": None,
                        "scope": "regular",
                        "name": sub.get("name"),
                        "ref_number": sub.get("ref_number"),
                        "template": sub.get("template"),
                        "sort_order": sub_idx,
                    }
                )

                for inp_idx, inp in enumerate(sub.get("inputs", []), 1):
                    seed["template_inputs"].append(
                        {
                            "input_uid": f"inp_{sub_uid}_{inp_idx}",
                            "owner_scope": "subcategory",
                            "owner_uid": sub_uid,
                            "name": inp.get("name"),
                            "label": inp.get("label"),
                            "sort_order": inp_idx,
                        }
                    )

        special_subs = category.get("special_subcategories", [])
        if isinstance(special_subs, list):
            for sub_idx, sub in enumerate(special_subs, 1):
                sub_uid = f"sub_special_{cat_uid}_{sub_idx}_{_slug(sub.get('name'))}"
                seed["subcategory_templates"].append(
                    {
                        "subcategory_uid": sub_uid,
                        "category_uid": cat_uid,
                        "wiring_uid": None,
                        "scope": "special",
                        "name": sub.get("name"),
                        "ref_number": sub.get("ref_number"),
                        "template": sub.get("template"),
                        "sort_order": sub_idx,
                    }
                )

                for inp_idx, inp in enumerate(sub.get("inputs", []), 1):
                    seed["template_inputs"].append(
                        {
                            "input_uid": f"inp_{sub_uid}_{inp_idx}",
                            "owner_scope": "subcategory",
                            "owner_uid": sub_uid,
                            "name": inp.get("name"),
                            "label": inp.get("label"),
                            "sort_order": inp_idx,
                        }
                    )

    return seed


def serialize_catalog(categories: List[Dict[str, Any]], source: str = "manager") -> Dict[str, Any]:
    return {
        "schema": {
            "name": "inspection_categories",
            "version": 2,
            "postgres_ready": True,
        },
        "metadata": {
            "updated_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            "source": source,
        },
        "categories": categories,
        "postgres_seed": _build_postgres_seed(categories),
    }
