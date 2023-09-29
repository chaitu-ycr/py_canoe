# Import Python Libraries here
import pythoncom
import win32com.client
from typing import Union

class Networks:
    def __init__(self, app_obj) -> None:
        self.app_obj = app_obj
        self.log = self.app_obj.log
        self.networks_com_obj = win32com.client.Dispatch(self.app_obj.app_com_obj.Networks)
    
    def fetch_diag_devices(self) -> dict:
        diag_devices = {}
        for network in self.networks_com_obj:
            for device in network.Devices:
                try:
                    diag_devices[device.Name] = device.Diagnostic
                except pythoncom.com_error:
                    pass
        return diag_devices

class Network:
    def __init__(self, network_com_obj) -> None:
        self.network_com_obj = network_com_obj
    
    @property
    def bus_type(self) -> int:
        """Determines the bus type of the network

        Returns:
            int: The type of the network: 0-Invalid, 1-CAN, 2-J1939, 5-LIN, 6-MOST, 7-FlexRay, 9-J1708, 11-Ethernet
        """
        return self.network_com_obj.BusType

    @property
    def devices(self) -> object:
        """Returns the Devices object.

        Returns:
            object: The Devices object
        """
        return self.network_com_obj.Devices
    
    @property
    def name(self) -> str:
        return self.network_com_obj.Name