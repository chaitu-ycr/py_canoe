import win32com.client

from py_canoe.core.child_elements.data_source import DataSource


class DataSourceIssue:
    """
    Represents a single issue from a DataSource import operation.
    """
    def __init__(self, com_object) -> None:
        self.com_object = win32com.client.Dispatch(com_object)

    @property
    def description(self) -> str:
        return self.com_object.Description

    @property
    def emitter(self) -> 'DataSource':
        return DataSource(self.com_object.Emitter)

    @property
    def severity(self) -> int:
        return self.com_object.Severity
