# -*- coding: utf-8 -*-
"""Обновление ключей спецификаций по источнику правил ГЭСН."""

from __future__ import absolute_import

import os
import sys

from pyrevit import forms

THIS_DIR = os.path.dirname(__file__)
BASE_DIR = os.path.dirname(THIS_DIR)
LIB_DIR = os.path.join(BASE_DIR, "lib")
if BASE_DIR not in sys.path:
    sys.path.append(BASE_DIR)
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

from lib import config, gesn_rules, spec_keys_cache  # noqa: E402

SOURCE_EXCEL = "excel"
SOURCE_DB = "db"

FSB_CODE_COLUMNS = [
    u"Код ФСБЦ 04.3.01.12 Растворы цементно-известковые",
    u"Код ФСБЦ 06.1.01.02 Камни керамические одинарные",
    u"Код ФСБЦ 06.1.01.05 Кирпичи керамические",
]


try:
    STRING_TYPES = (basestring,)  # type: ignore[name-defined]
except NameError:  # pragma: no cover
    STRING_TYPES = (str,)

try:
    unicode  # type: ignore[name-defined]
except NameError:  # pragma: no cover
    unicode = str  # type: ignore[assignment]


def _t(value):
    try:
        return unicode(value)  # type: ignore[name-defined]
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def _sorted_values(values):
    normalized = {_t(value) for value in values if value not in (None, u"", "")}
    return sorted(normalized, key=lambda x: x.lower())


def _collect_rule_based_values(rules):
    unique = {
        u"Стадия": set(),
        u"Армирование": set(),
        u"Размеры кирпича": set(),
    }

    for rule in rules or []:
        stage = (getattr(rule, "stage", u"") or u"").strip()
        if stage:
            stage_display = stage.capitalize() if stage.islower() else stage
            unique[u"Стадия"].add(stage_display)
        reinforcement = getattr(rule, "reinforcement", u"") or u""
        if reinforcement:
            display_reinf = reinforcement
            if reinforcement in {u"да", u"нет"}:
                display_reinf = reinforcement.capitalize()
            unique[u"Армирование"].add(display_reinf)
        brick = getattr(rule, "brick_size", u"") or u""
        if brick:
            unique[u"Размеры кирпича"].add(brick)

    return {
        key: _sorted_values(values)
        for key, values in unique.items()
        if values
    }


def _collect_fsb_values_from_excel(excel_path):
    try:
        collected = gesn_rules.collect_column_values_from_excel(
            path=excel_path,
            sheet_name=getattr(config, "EXCEL_SHEET_NAME", None),
            columns=FSB_CODE_COLUMNS,
        )
    except Exception:
        return {}

    normalized = {}
    for column, values in collected.items():
        if not values:
            continue
        normalized[column] = _sorted_values(values)
    return normalized


def _save_result(source_type, excel_path, rules, unique_values):
    spec_keys_cache.save_cache(
        source_type=source_type,
        excel_path=excel_path,
        rules=rules,
        unique_values=unique_values,
    )


def _handle_excel_source():
    excel_path = forms.pick_file(
        file_ext="xlsx",
        multi_file=False,
        title=u"Выберите файл Excel с таблицей ГЭСН",
    )
    if not excel_path:
        forms.alert(u"Файл не выбран. Операция отменена.", exitscript=True)
        return

    try:
        rules = gesn_rules.load_rules_from_excel(path=excel_path)
    except Exception as exc:
        forms.alert(
            u"Не удалось загрузить правила из файла:\n{0}\n\nОшибка: {1}".format(
                excel_path,
                exc,
            ),
            exitscript=True,
        )
        return

    unique_values = _collect_rule_based_values(rules)
    unique_values.update(_collect_fsb_values_from_excel(excel_path))

    _save_result(SOURCE_EXCEL, excel_path, rules, unique_values)

    forms.alert(
        u"Ключи спецификаций обновлены из файла:\n{0}\nЗагружено правил: {1}".format(
            excel_path,
            len(rules),
        )
    )


def _handle_db_source():
    try:
        rules = gesn_rules.load_rules_from_db()
    except NotImplementedError as exc:
        forms.alert(_t(exc), exitscript=True)
        return
    except Exception as exc:
        forms.alert(
            u"Не удалось загрузить правила из базы данных: {0}".format(exc),
            exitscript=True,
        )
        return

    unique_values = _collect_rule_based_values(rules)
    _save_result(SOURCE_DB, None, rules, unique_values)

    forms.alert(
        u"Ключи спецификаций обновлены из базы данных. Загружено правил: {0}".format(
            len(rules),
        )
    )


def main():
    option_map = {
        u"Из файла Excel": SOURCE_EXCEL,
        u"Из базы данных": SOURCE_DB,
    }
    selected = forms.alert(
        u"Выберите источник правил ГЭСН",
        options=list(option_map.keys()),
        title=u"Обновить ключи спецификаций",
    )

    if not selected:
        forms.alert(u"Операция отменена пользователем.", exitscript=True)
        return

    source = option_map.get(selected, SOURCE_EXCEL)

    if source == SOURCE_DB:
        _handle_db_source()
    else:
        _handle_excel_source()


if __name__ == "__main__":
    main()
