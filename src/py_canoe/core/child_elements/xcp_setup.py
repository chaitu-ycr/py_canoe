import win32com.client

from py_canoe.core.child_elements.mc_ecus import McECUs


class XCPSetup:
    def __init__(self, xcp_setup_com_object) -> None:
        self.com_object = win32com.client.Dispatch(xcp_setup_com_object)

    @property
    def mcus(self) -> 'McECUs':
        return McECUs(self.com_object.MCUs)
