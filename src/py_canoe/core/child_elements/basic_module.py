import win32com.client


class BasicModule:
    """
    Represents a basic VT System module (e.g., VT1004, VT2004, etc.).
    """
    def __init__(self, basic_module_com_obj):
        self.com_object = win32com.client.Dispatch(basic_module_com_obj)

    @property
    def name(self) -> str:
        """The name of the module."""
        return self.com_object.Name

    @property
    def type(self) -> int:
        """The type of the module (e.g. 1004)."""
        return self.com_object.Type
