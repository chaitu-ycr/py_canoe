from py_canoe.core.child_elements.variables_file import VariablesFile


class VariablesFiles:
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def count(self):
        return self.com_object.Count

    def item(self, index: int) -> VariablesFile:
        return VariablesFile(self.com_object.Item(index))

    def add(self, name: str):
        return self.com_object.Add(name)

    def remove(self, variable_file):
        return self.com_object.Remove(variable_file)
