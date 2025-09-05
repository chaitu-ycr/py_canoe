import win32com.client

from py_canoe.core.child_elements.test_modules import TestModules
from py_canoe.core.child_elements.test_setup_folders import TestSetupFolders


class TestSetupFolderExt:
    """The TestSetupFolderExt object represents a directory in CANoe's test setup."""
    def __init__(self, com_object) -> None:
        self.com_object = win32com.client.Dispatch(com_object)

    @property
    def enabled(self) -> bool:
        return self.com_object.Enabled

    @enabled.setter
    def enabled(self, enabled: bool):
        self.com_object.Enabled = enabled

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def folders(self) -> 'TestSetupFolders':
        return TestSetupFolders(self.com_object.Folders)

    @property
    def test_modules(self) -> 'TestModules':
        return TestModules(self.com_object.TestModules)

    def execute_all(self):
        self.com_object.ExecuteAll()

    def stop_sequence(self):
        self.com_object.StopSequence()
