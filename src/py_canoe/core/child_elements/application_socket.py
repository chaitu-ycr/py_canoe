class ApplicationSocket:
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def bus_registry(self) -> bytearray:
        return self.com_object.BusRegistry

    @property
    def fb_lock_list(self) -> bytearray:
        return self.com_object.FBlockList
