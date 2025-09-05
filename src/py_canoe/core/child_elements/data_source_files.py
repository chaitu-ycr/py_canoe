from py_canoe.core.child_elements.data_source_file import DataSourceFile


class DataSourceFiles:
    """
    Collection of DataSourceFile objects for a FileGroupDataSource.
    """
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'DataSourceFile':
        return DataSourceFile(self.com_object.Item(index))

    def add(self, source_file_path: str) -> 'DataSourceFile':
        return DataSourceFile(self.com_object.Add(source_file_path))

    def remove(self, index: int):
        self.com_object.Remove(index)
