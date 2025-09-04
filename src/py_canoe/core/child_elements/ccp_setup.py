import win32com.client

from py_canoe.core.child_elements.mc_ecus import McECUs


class CCPSetup:
    """
    The CCPSetup object represents the CCP settings of a CANoe configuration.
    """
    def __init__(self, ccp_setup_com_object) -> None:
        self.com_object = win32com.client.Dispatch(ccp_setup_com_object)

    @property
    def ecus(self) -> 'McECUs':
        return McECUs(self.com_object.ECUs)
