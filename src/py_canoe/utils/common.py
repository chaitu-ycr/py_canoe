import logging

def create_console_handler(level=logging.DEBUG, fmt="%(asctime)s [PY_CANOE] [%(levelname)-4.8s] %(message)s"):
    """Create a console logging handler."""
    handler = logging.StreamHandler()
    handler.setLevel(level)
    handler.setFormatter(logging.Formatter(fmt))
    return handler

def create_file_handler(filename='py_canoe.log', level=logging.DEBUG, fmt="%(asctime)s [PY_CANOE] [%(levelname)-4.8s] %(message)s"):
    """Create a file logging handler."""
    handler = logging.FileHandler(filename, encoding='utf-8')
    handler.setLevel(level)
    handler.setFormatter(logging.Formatter(fmt))
    return handler

def setup_logger(name='py_canoe', filename='py_canoe.log'):
    """Set up and return a logger with console and file handlers."""
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    if not logger.handlers:
        logger.addHandler(create_console_handler())
        logger.addHandler(create_file_handler(filename=filename))
    return logger

logger = setup_logger()
