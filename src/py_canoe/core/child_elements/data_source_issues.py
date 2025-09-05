from py_canoe.core.child_elements.data_source_issue import DataSourceIssue


class DataSourceIssues:
    """
    Collection of DataSourceIssue objects from the last import operation.
    """
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'DataSourceIssue':
        return DataSourceIssue(self.com_object.Item(index))
