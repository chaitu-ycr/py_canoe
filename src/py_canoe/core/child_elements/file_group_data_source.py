from py_canoe.core.child_elements.data_source_files import DataSourceFiles


class FileGroupDataSource:
    """
    Represents a group data source that can import multiple files simultaneously.
    """
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def source_files(self) -> 'DataSourceFiles':
        return DataSourceFiles(self.com_object.SourceFiles)
