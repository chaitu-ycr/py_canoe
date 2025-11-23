import win32com.client


class ApplicationSpecificModule:
    """
    Represents an application specific VT System module (e.g., VT7900 with application board or UserFPGA).
    """
    def __init__(self, app_specific_module_com_obj):
        self.com_object = win32com.client.Dispatch(app_specific_module_com_obj)

    @property
    def name(self) -> str:
        """The name of the module. For User FPGA, includes project name."""
        return self.com_object.Name

    @property
    def base_type(self) -> int:
        """The basic module type (e.g. 1004)."""
        return self.com_object.BaseType

    @property
    def id(self) -> int:
        """The unique ID of this application specific module."""
        return self.com_object.ID
