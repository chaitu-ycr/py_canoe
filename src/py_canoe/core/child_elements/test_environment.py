import win32com.client

from py_canoe.core.child_elements.test_modules import TestModules
from py_canoe.core.child_elements.test_setup_folders import TestSetupFolders


class TestEnvironment:
    """The TestEnvironment object represents a test environment within CANoe's test setup."""
    def __init__(self, com_object):
        self.com_object = win32com.client.Dispatch(com_object)
        self.__test_modules = TestModules(self.com_object.TestModules)
        self.__test_setup_folders = TestSetupFolders(self.com_object.Folders)
        self.__all_test_modules = {}
        self.__all_test_setup_folders = {}

    @property
    def enabled(self) -> bool:
        return self.com_object.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.com_object.Enabled = value

    @property
    def full_name(self) -> str:
        return self.com_object.FullName

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def path(self) -> str:
        return self.com_object.Path

    def execute_all(self) -> None:
        self.com_object.ExecuteAll()

    def save(self, name: str, prompt_user=False) -> None:
        self.com_object.Save(name, prompt_user)

    def save_as(self, name: str, major: int, minor: int, prompt_user=False) -> None:
        self.com_object.SaveAs(name, major, minor, prompt_user)

    def stop_sequence(self) -> None:
        self.com_object.StopSequence()

    def update_all_test_setup_folders(self, tsfs_instance=None):
        if tsfs_instance is None:
            tsfs_instance = self.__test_setup_folders
        if tsfs_instance.count > 0:
            test_setup_folders = tsfs_instance.fetch_test_setup_folders()
            for tsf_name, tsf_inst in test_setup_folders.items():
                self.__all_test_setup_folders[tsf_name] = tsf_inst
                if tsf_inst.test_modules.count > 0:
                    self.__all_test_modules.update(tsf_inst.test_modules.fetch_test_modules())
                if tsf_inst.folders.count > 0:
                    tsfs_instance = tsf_inst.folders
                    self.update_all_test_setup_folders(tsfs_instance)

    def get_all_test_modules(self):
        self.update_all_test_setup_folders()
        self.__all_test_modules.update(self.__test_modules.fetch_test_modules())
        return self.__all_test_modules
