class Encoding:
    def __init__(self, encoding):
        self.com_object = encoding

    @property
    def factor(self) -> float:
        return self.com_object.Factor

    @property
    def lower_bound(self) -> int:
        return self.com_object.LowerBound

    @property
    def offset(self) -> float:
        return self.com_object.Offset

    @property
    def text(self) -> str:
        return self.com_object.Text

    @property
    def unit(self) -> str:
        return self.com_object.Unit

    @property
    def upper_bound(self) -> int:
        return self.com_object.UpperBound
