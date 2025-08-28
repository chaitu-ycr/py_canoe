import win32com.client


class Databases:
    """The Databases object represents the assigned databases of CANoe."""
    def __init__(self, databases_com_obj):
        self.com_object = win32com.client.Dispatch(databases_com_obj)

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'Database':
        return Database(self.com_object.Item(index))

    def add(self, full_name: str) -> 'Database':
        return Database(self.com_object.Add(full_name))

    def add_network(self, database_name: str, network_name: str) -> 'Database':
        return Database(self.com_object.AddNetwork(database_name, network_name))

    def remove(self, index: int) -> None:
        self.com_object.Remove(index)


class Database:
    """The Database object represents the assigned database of the CANoe application."""
    def __init__(self, database_com_obj):
        self.com_object = win32com.client.Dispatch(database_com_obj)

    @property
    def channel(self) -> int:
        return self.com_object.Channel

    @channel.setter
    def channel(self, channel: int) -> None:
        self.com_object.Channel = channel

    @property
    def full_name(self) -> str:
        return self.com_object.FullName

    @full_name.setter
    def full_name(self, full_name: str) -> None:
        self.com_object.FullName = full_name

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def path(self) -> str:
        return self.com_object.Path
