import os
import sys
import logging
from logging.handlers import RotatingFileHandler


def get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    # 2 niveles hacia arriba desde src/utils/
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


LOG_FILE = os.path.join(get_base_dir(), "app.log")


def setup_logger():
    logger = logging.getLogger("JiraTeamsSync")
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger

    file_handler = RotatingFileHandler(
        LOG_FILE, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    if sys.stdout is not None:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    return logger


logger = setup_logger()
