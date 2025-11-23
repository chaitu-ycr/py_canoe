from py_canoe.core.child_elements.data_source import DataSource
from py_canoe.core.child_elements.data_sources import DataSources


class DataSourceSetup:
    """
    Provides the data source management API.
    """
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def data_sources(self) -> 'DataSources':
        return DataSources(self.com_object.DataSources)

    def get_data_source_by_id(self, data_source_id: int) -> 'DataSource':
        return DataSource(self.com_object.GetDataSourceById(data_source_id))
