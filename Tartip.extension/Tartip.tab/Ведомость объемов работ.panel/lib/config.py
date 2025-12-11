# -*- coding: utf-8 -*-
"""
Константы конфигурации для работы с таблицей ГЭСН.
"""
import os

# Excel-файл с правилами сопоставления.
EXCEL_FILE_NAME = u"БД по стене.xlsx"
# Имя листа с правилами.
EXCEL_SHEET_NAME = u"ГЭСН_стены_перегородки"

# Толщина сопоставляется с допуском в миллиметрах.
THICKNESS_TOLERANCE_MM = 0.5

# Параметр объёма по умолчанию, если в таблице не указано иное.
DEFAULT_VOLUME_PARAM = u"Площадь"

# Поведение при отсутствии правил: если True, параметр будет очищен.
CLEAR_CODE_WHEN_MISS = False

# Путь к файлу Excel рядом с расширением.
BASE_DIR = os.path.dirname(os.path.dirname(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, EXCEL_FILE_NAME)
