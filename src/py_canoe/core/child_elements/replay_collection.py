import win32com.client


class ReplayCollection:
    """The ReplayCollection object represents the Replay Blocks of the CANoe application."""
    def __init__(self, replay_collection_com_obj):
        self.com_object = replay_collection_com_obj

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'ReplayBlock':
        return ReplayBlock(self.com_object.Item(index))

    def add(self, name: str) -> 'ReplayBlock':
        return ReplayBlock(self.com_object.Add(name))

    def remove(self, index: int):
        self.com_object.Remove(index)


class ReplayBlock:
    """The ReplayBlock object represents a Replay Block of the CANoe application."""
    def __init__(self, replay_block_com_obj):
        self.com_object = win32com.client.Dispatch(replay_block_com_obj)

    @property
    def enabled(self) -> bool:
        return self.com_object.Enabled

    @enabled.setter
    def enabled(self, value: bool):
        self.com_object.Enabled = value

    @property
    def name(self) -> str:
        return self.com_object.Name

    @name.setter
    def name(self, value: str):
        self.com_object.Name = value

    @property
    def path(self) -> str:
        return self.com_object.Path

    @path.setter
    def path(self, value: str):
        self.com_object.Path = value

    def start(self):
        self.com_object.Start()

    def stop(self):
        self.com_object.Stop()
