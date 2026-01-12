from py_canoe.core.child_elements.test_setup_folder_ext import TestSetupFolderExt


class TestSetupFolders:
    """The TestSetupFolders object represents the folders in a test environment or in a test setup folder."""
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def count(self) -> int:
        return self.com_object.Count

    def add(self, full_name: str) -> 'TestSetupFolderExt':
        return TestSetupFolderExt(self.com_object.Add(full_name))

    def remove(self, index: int, prompt_user=False) -> None:
        self.com_object.Remove(index, prompt_user)

    def fetch_test_setup_folders(self) -> dict:
        test_setup_folders = dict()
        for index in range(1, self.count + 1):
            tsf_com_obj = self.com_object.Item(index)
            tsf_inst = TestSetupFolderExt(tsf_com_obj)
            test_setup_folders[tsf_inst.name] = tsf_inst
        return test_setup_folders
