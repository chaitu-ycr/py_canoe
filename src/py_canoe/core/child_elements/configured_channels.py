from py_canoe.core.child_elements.configured_channel import ConfiguredChannel


class ConfiguredChannels:
    """
    Python wrapper for CANoe COM ConfiguredChannels object.
    Represents the collection of channels on a configured VT System module.
    """
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def count(self):
        """Returns the number of channels in the collection."""
        return self.com_object.Count

    def item(self, index) -> ConfiguredChannel:
        """
        Returns the ConfiguredChannel object at the given index (1-based).
        """
        return ConfiguredChannel(self.com_object.Item(index))
