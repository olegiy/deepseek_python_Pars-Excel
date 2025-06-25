# backup_utils.py

import os
import shutil
from datetime import datetime
import logging

logger = logging.getLogger(__name__)

def create_backup(file_path):
    """Создаёт резервную копию файла."""
    backup_dir = os.path.join(os.path.dirname(file_path), "backups")
    os.makedirs(backup_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{os.path.splitext(os.path.basename(file_path))[0]}_backup_{timestamp}.xlsx"
    backup_path = os.path.join(backup_dir, backup_name)
    try:
        shutil.copy2(file_path, backup_path)
        logger.info(f"✅ Создана резервная копия: {backup_path}")
        return backup_path
    except Exception as e:
        logger.error(f"❌ Ошибка при создании резервной копии: {e}")
        return None