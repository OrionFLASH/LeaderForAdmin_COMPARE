import logging
import os
from datetime import datetime

def setup_logger(log_dir, basename):
    """
    Создаёт логгер, который пишет и в файл (append, по дате), и в консоль.
    log_dir: путь к папке для логов
    basename: имя (без даты) для лог-файла
    """
    now = datetime.now()
    day_str = now.strftime("%Y%m%d")
    time_str = now.strftime("%H:%M:%S")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, f"{basename}_{day_str}.log")
    # Добавляем маркер начала новой сессии
    with open(log_path, "a", encoding="utf-8") as logf:
        logf.write(f"\n-------- NEW LOG START AT {day_str} ({time_str}) -------\n")
    # Стандартное подключение логгера
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    if logger.hasHandlers():
        logger.handlers.clear()
    fh = logging.FileHandler(log_path, encoding='utf-8', mode='a')
    fh.setLevel(logging.DEBUG)
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    fmt = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s', "%Y-%m-%d %H:%M:%S")
    fh.setFormatter(fmt)
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    logging.info(f"Лог-файл активен (append): {log_path}")
    return logger
