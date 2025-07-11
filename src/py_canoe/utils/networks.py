# TODO: Implement Networks
import logging
import win32com.client

logging.getLogger('py_canoe')

class Networks:
    def __init__(self, app):
        self.com_object = win32com.client.Dispatch(app.com_object.Networks)