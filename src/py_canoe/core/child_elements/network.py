from py_canoe.core.child_elements.devices import Devices


class Network:
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def bus_type(self) -> int:
        return self.com_object.BusType

    @property
    def devices(self) -> Devices:
        return Devices(self.com_object.Devices)

    @property
    def name(self) -> str:
        return self.com_object.Name
