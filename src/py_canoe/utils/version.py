import logging
import win32com.client

logging.getLogger('py_canoe')

class Version:
    def __init__(self, application):
        self.com_object = win32com.client.Dispatch(application.com_object.Version)

    def __str__(self):
        return f"{self.full_name}"

    @property
    def build(self):
        return self.com_object.Build

    @property
    def full_name(self):
        return self.com_object.FullName

    @property
    def major(self):
        return self.com_object.major

    @property
    def minor(self):
        return self.com_object.minor

    @property
    def name(self):
        return self.com_object.Name

    @property
    def patch(self):
        return self.com_object.Patch
