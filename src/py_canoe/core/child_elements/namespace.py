from py_canoe.core.child_elements.variables import Variables
from py_canoe.core.child_elements.namespaces import Namespaces


class Namespace:
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def comment(self):
        return self.com_object.Comment

    @property
    def name(self):
        return self.com_object.Name

    def variables(self) -> Variables:
        return Variables(self.com_object.Variables)

    def namespaces(self) -> 'Namespaces':
        return Namespaces(self.com_object.Namespaces)
