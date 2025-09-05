from py_canoe.core.child_elements.application_model_file import ApplicationModelFile


class ApplicationModelFiles:
    """
    Collection of ApplicationModelFile objects (immutable).
    """
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'ApplicationModelFile':
        return ApplicationModelFile(self.com_object.Item(index))

    def add(self, application_model_file_path: str) -> 'ApplicationModelFile':
       self.com_object.Add(application_model_file_path)

    def remove(self, index: int):
        self.com_object.Remove(index)
