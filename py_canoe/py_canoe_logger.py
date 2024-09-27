# Import Python Libraries here
import os
import sys
import logging
from logging import handlers


class PyCanoeLogger:
    """
    PyCanoeLogger is a class that provides logging functionality for the PyCanoe application.
    Args:
        py_canoe_log_dir (str): The directory path where the log files will be stored. Defaults to an empty string.
    """

    def __init__(self, py_canoe_log_dir='') -> None:
        self.log = logging.getLogger('CANOE_LOG')
        self.log.handlers.clear()
        self.log.propagate = False
        self.__py_canoe_log_initialization(py_canoe_log_dir)

    def __py_canoe_log_initialization(self, py_canoe_log_dir):
        self.log.setLevel(logging.DEBUG)
        log_format = logging.Formatter("%(asctime)s [CANOE_LOG] [%(levelname)-4.8s] %(message)s")
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(log_format)
        self.log.addHandler(ch)
        if py_canoe_log_dir != '' and not os.path.exists(py_canoe_log_dir):
            os.makedirs(py_canoe_log_dir, exist_ok=True)
        if os.path.exists(py_canoe_log_dir):
            fh = handlers.RotatingFileHandler(fr'{py_canoe_log_dir}\py_canoe.log', maxBytes=0, encoding='utf-8')
            fh.setFormatter(log_format)
            self.log.addHandler(fh)
