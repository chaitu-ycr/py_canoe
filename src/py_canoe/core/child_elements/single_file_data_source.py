from py_canoe.core.child_elements.data_source_file import DataSourceFile


class SingleFileDataSource:
    """
    Represents a data source for a single file.
    """
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def file(self) -> 'DataSourceFile':
        return DataSourceFile(self.com_object.File)
