import win32com.client

from py_canoe.core.child_elements.application_specific_modules import ApplicationSpecificModules
from py_canoe.core.child_elements.basic_modules import BasicModules


class AvailableModules:
    """
    Represents the collection of known VT module types that can be added to VT System.
    """
    def __init__(self, available_modules_com_obj):
        self.com_object = win32com.client.Dispatch(available_modules_com_obj)

    @property
    def application_specific_modules(self) -> 'ApplicationSpecificModules':
        """Returns the collection of available application specific modules."""
        return ApplicationSpecificModules(self.com_object.ApplicationSpecificModules)

    @property
    def basic_modules(self) -> 'BasicModules':
        """Returns the collection of available basic modules."""
        return BasicModules(self.com_object.BasicModules)
