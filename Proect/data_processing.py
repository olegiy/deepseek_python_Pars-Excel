# data_processing.py

import math
import re
import logging

logger = logging.getLogger(__name__)

def try_convert(value):
    """Преобразует значение в число."""
    if value is None or value == "ERROR:#VALUE!":
        logger.debug(f"🚫 Значение пропущено: {value}")
        return None, None
    if isinstance(value, (int, float)):
        converted = math.ceil(value) if value != int(value) else int(value)
        logger.debug(f"🔢 Преобразовано число: {value} → {converted}")
        return converted, '0'
    if isinstance(value, str):
        value = value.strip()
        if '/' in value and value.replace('/', '', 1).isdigit():
            converted = int(value.split('/')[0])
            logger.debug(f"🔢 Преобразовано дробное значение: {value} → {converted}")
            return converted, '0'
        try:
            num = float(value.replace(',', '.'))
            converted = math.ceil(num) if num != int(num) else int(num)
            logger.debug(f"🔢 Преобразовано значение: {value} → {converted}")
            return converted, '0'
        except ValueError:
            logger.debug(f"🔤 Не удалось преобразовать строку: {value}")
            return value, None
    logger.debug(f"🔁 Без изменений: {value}")
    return value, None

def extract_thickness_value(text, keyword):
    """Извлекает значение толщины из текста."""
    try:
        pattern = re.compile(rf"{re.escape(keyword)}[^\d]*([\d,\.]+)")
        matches = pattern.findall(text.lower())
        if matches:
            value = float(matches[-1].replace(',', '.'))
            logger.debug(f"📐 Извлечена толщина: {value}")
            return value
    except Exception as e:
        logger.error(f"❌ Ошибка при извлечении толщины: {e}")
    return None

def get_section_name(cell_value):
    """Извлекает название секции."""
    if not isinstance(cell_value, str):
        return None
    match = re.search(r'section:\s*(.+?)(?:\s*толщина|thickness|$)', cell_value.lower())
    if match:
        logger.debug(f"🔖 Извлечено имя секции: {match.group(1).strip()}")
        return match.group(1).strip()
    return None