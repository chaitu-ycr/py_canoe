from py_canoe.core.child_elements.application_model import ApplicationModel


class ApplicationModels:
    """
    Collection of ApplicationModel objects.
    """
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'ApplicationModel':
        return ApplicationModel(self.com_object.Item(index))

    def add(self, file_path: str) -> 'ApplicationModel':
        return ApplicationModel(self.com_object.Add(file_path))

    def remove(self, index: int):
        self.com_object.Remove(index)
