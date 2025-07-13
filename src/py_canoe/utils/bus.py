# TODO: complete implementation of the Bus class
import logging
import win32com.client

logging.getLogger('py_canoe')

class Bus:
    def __init__(self, application, bus_type: str = 'CAN'):
        self.bus_type = bus_type
        self.com_object = win32com.client.Dispatch(application.com_object.Bus(bus_type))

    def active(self) -> bool:
        return self.com_object.Active

    def baudrate(self) -> int:
        return self.com_object.Baudrate

    def name(self) -> str:
        return self.com_object.Name

    def set_name(self, name: str):
        self.com_object.Name = name

    def get_signal(self, channel: str, message: str, signal: str):
        return self.com_object.GetSignal(channel, message, signal)

    def get_j1939_signal(self, channel: str, message: str, signal: str, source_address: str, destination_address: str):
        return self.com_object.GetJ1939Signal(channel, message, signal, source_address, destination_address)
