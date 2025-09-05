import win32com.client


class Participant:
    """
    Represents a single participant associated with an ApplicationModel.
    """
    def __init__(self, com_object) -> None:
        self.com_object = win32com.client.Dispatch(com_object)

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def path(self) -> str:
        return self.com_object.Path
