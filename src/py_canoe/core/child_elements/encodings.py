from py_canoe.core.child_elements.encoding import Encoding


class Encodings:
    def __init__(self, encodings):
        self.com_object = encodings

    @property
    def count(self):
        return self.com_object.Count

    def item(self, index: int) -> 'Encoding':
        return Encoding(self.com_object.Item(index))
