import win32com.client


class ConnectedModules:
    """Represents the collection of VT System modules currently connected to the computer."""
    def __init__(self, connected_modules_com_obj):
        self.com_object = win32com.client.Dispatch(connected_modules_com_obj)

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int):
        return self.com_object.Item(index)  # Returns a ConnectedModule object
