# -*- coding: utf-8 -*-
"""Кэш конфигурации источника правил ГЭСН и ключевых значений."""
from __future__ import absolute_import

import io
import json
import os

THIS_DIR = os.path.dirname(__file__)
CACHE_FILE = os.path.join(THIS_DIR, "spec_keys_cache.json")


def _serialize_unique_values(unique_values):
    if not unique_values:
        return {}

    serialized = {}
    for key, values in unique_values.items():
        if not values:
            continue
        normalized = sorted({value for value in values if value not in (None, "")})
        if normalized:
            serialized[key] = normalized
    return serialized


def save_cache(source_type, excel_path=None, rules=None, unique_values=None):
    """Сохраняет сведения об источнике и агрегированные ключевые значения."""

    data = {
        "source_type": source_type,
        "excel_path": excel_path,
        "rules_count": len(rules) if rules is not None else None,
        "unique_values": _serialize_unique_values(unique_values),
    }

    cache_dir = os.path.dirname(CACHE_FILE)
    if cache_dir and not os.path.exists(cache_dir):
        os.makedirs(cache_dir)

    with io.open(CACHE_FILE, "w", encoding="utf-8") as fp:
        json.dump(data, fp, ensure_ascii=False, indent=2)

    return data


def load_cache():
    """Возвращает сохранённые параметры источника или None."""

    if not os.path.exists(CACHE_FILE):
        return None

    with io.open(CACHE_FILE, "r", encoding="utf-8") as fp:
        return json.load(fp)
