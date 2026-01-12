import win32com.client

from py_canoe.core.child_elements.basic_module import BasicModule


class BasicModules:
    """Represents the collection of available basic modules."""
    def __init__(self, basic_modules_com_obj):
        self.com_object = win32com.client.Dispatch(basic_modules_com_obj)

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'BasicModule':
        return BasicModule(self.com_object.Item(index))
