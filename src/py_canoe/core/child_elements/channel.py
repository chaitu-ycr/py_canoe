import win32com.client

from py_canoe.core.child_elements.can_controller import CanController


class Channel:
    def __init__(self, com_object):
        self.com_object = win32com.client.Dispatch(com_object)

    @property
    def bus_type(self) -> int:
        return self.com_object.BusType

    @property
    def controller(self) -> 'CanController':
        return CanController(self.com_object.Controller)

    @property
    def number(self) -> int:
        return self.com_object.Number
