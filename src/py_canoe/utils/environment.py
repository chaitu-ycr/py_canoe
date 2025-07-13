# TODO: complete implementation of the Environment class
import logging
import win32com.client

logging.getLogger('py_canoe')

class Environment:
    def __init__(self, app):
        self.com_object = win32com.client.Dispatch(app.com_object.Environment)
