from py_canoe.core.child_elements.application_models import ApplicationModels
from py_canoe.core.child_elements.data_source import DataSource


class VttSutImportResult:
    """
    The VttSutImportResult object encapsulates the results of the import vVIRTUALtarget SUTs.
    """
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def created_application_models(self) -> 'ApplicationModels':
        return ApplicationModels(self.com_object.CreatedApplicationModels)

    @property
    def created_data_source(self) -> 'DataSource':
        return DataSource(self.com_object.CreatedDataSource)

    def success(self) -> bool:
        return self.com_object.Success
