# import os
import sys
import time
import logging
import pythoncom
from datetime import datetime

def setup_logger(name='py_canoe', filename='py_canoe.log'):
    """Set up and return a logger with console and file handlers."""
    # os.makedirs(os.path.dirname(filename), exist_ok=True)
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    fmt = "%(asctime)s [PY_CANOE] [%(levelname)-4.8s] %(message)s"
    if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(logging.Formatter(fmt))
        logger.addHandler(console_handler)
    if not any(isinstance(h, logging.FileHandler) for h in logger.handlers):
        file_handler = logging.FileHandler(filename, mode='w', encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(logging.Formatter(fmt))
        logger.addHandler(file_handler)
    logger.propagate = False
    return logger

logger = setup_logger()

def wait(timeout_seconds=0.1):
    """Pump waiting COM messages and sleep for the given timeout."""
    pythoncom.PumpWaitingMessages()
    time.sleep(timeout_seconds)

def DoEvents() -> None:
    wait(0.01)

def DoEventsUntil(cond, timeout, title) -> bool:
    base_time = datetime.now()
    while not cond():
        DoEvents()
        now = datetime.now()
        difference = now - base_time
        seconds = difference.seconds
        if seconds > timeout:
            logger.warning(f'⚠️ {title} timeout({timeout} s)')
            return False
    return True
