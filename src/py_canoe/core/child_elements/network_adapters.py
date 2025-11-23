import win32com.client

class NetworkAdapters:
    """Represents the collection of available network adapters for VT System communication."""
    def __init__(self, network_adapters_com_obj):
        self.com_object = win32com.client.Dispatch(network_adapters_com_obj)

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int):
        return self.com_object.Item(index)  # Returns a NetworkAdapter object
