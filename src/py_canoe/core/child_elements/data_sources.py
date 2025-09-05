from py_canoe.core.child_elements.data_source import DataSource
from py_canoe.core.child_elements.file_group_data_source import FileGroupDataSource
from py_canoe.core.child_elements.single_file_data_source import SingleFileDataSource


class DataSources:
    """
    Collection of DataSource objects.
    """
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'DataSource':
        return DataSource(self.com_object.Item(index))

    def add_group_data_source(self, name: str) -> 'FileGroupDataSource':
        return FileGroupDataSource(self.com_object.AddGroupDataSource(name))

    def add_single_file_data_source(self, source_file_path: str) -> 'SingleFileDataSource':
        return SingleFileDataSource(self.com_object.AddSingleFileDataSource(source_file_path))

    def remove(self, index: int):
        self.com_object.Remove(index)
