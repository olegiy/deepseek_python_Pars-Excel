# file_utils.py

from tkinter import Tk, filedialog
import logging

logger = logging.getLogger(__name__)

def select_file():
    """Открывает окно выбора файла."""
    Tk().withdraw()
    logger.info("📂 Открыто окно выбора файла.")
    file_path = filedialog.askopenfilename(
        title="Выберите Excel-файл",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        logger.info(f"📁 Выбран файл: {file_path}")
    else:
        logger.warning("⚠️ Файл не выбран.")
    return file_path