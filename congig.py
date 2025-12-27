"""
Конфигурация сканера анкет Muzloto
"""

import os
from pathlib import Path

# Пути
BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
SCANS_DIR = BASE_DIR / "scans"

# Создаем необходимые папки
for folder in [DATA_DIR, SCANS_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

# Конфигурация Excel
EXCEL_CONFIG = {
    "file_name": "анкеты_muzloto.xlsx",
    "sheet_name": "Анкеты",
    "auto_format": True,
    "backup_count": 10,  # Количество бэкапов
}

# Конфигурация OCR
OCR_CONFIG = {
    "language": "rus",  # Язык распознавания
    "psm": 6,  # Page segmentation mode
    "oem": 3,  # OCR Engine mode
    "confidence_threshold": 70,  # Минимальная уверенность в %
}

# Поля анкеты Muzloto (для валидации)
MUZLOTO_FIELDS = {
    "required": ["Дата", "Место игры"],
    "optional": [
        "Номер столика",
        "Довольны ли вы посещением Музлого",
        "Понравился ли вам плейлист", 
        "Какие треки вы бы добавили",
        "Понравилась ли вам локация",
        "Понравилась ли вам кухня и бар",
        "Устроил ли вас сервис, время подачи",
        "Понравилась ли вам работа ведущего",
        "Сколько раз вы были на Музлого",
        "Оцените стоимость игры за билет",
        "Знаете ли вы, что Музлого можно заказать",
        "Откуда вы о нас узнали",
        "Ради чего вы обычно ходите на подобные вечеринки",
        "Что нам стоит улучшить",
        "Телефон"
    ]
}

# Валидация данных
VALIDATION = {
    "phone_patterns": [
        r"\+7\s?\d{3}\s?\d{3}\s?\d{2}\s?\d{2}",
        r"8\s?\d{3}\s?\d{3}\s?\d{2}\s?\d{2}",
        r"\d{10,11}"
    ],
    "date_patterns": [
        r"\d{1,2}\.\d{1,2}",  # 18.12
        r"\d{1,2}\.\d{1,2}\.\d{4}",  # 18.12.2023
    ]
}