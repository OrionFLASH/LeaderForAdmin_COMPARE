import logging
import os
from datetime import datetime
from config import LOG_BASENAME, LOG_DIR

def setup_logger():
    now = datetime.now()
    day_str = now.strftime("%Y%m%d")
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, f"{LOG_BASENAME}_{day_str}.log")
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
