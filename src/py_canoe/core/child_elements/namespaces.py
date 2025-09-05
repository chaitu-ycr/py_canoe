from py_canoe.core.child_elements.namespace import Namespace


class Namespaces:
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def count(self):
        return self.com_object.Count

    def item(self, index: int) -> Namespace:
        return Namespace(self.com_object.Item(index))

    def add(self, name: str):
        return self.com_object.Add(name)

    def remove(self, variable):
        return self.com_object.Remove(variable)
