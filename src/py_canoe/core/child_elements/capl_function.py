class CaplFunction:
    """
    The CAPLFunction object represents a CAPL function.
    """
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def parameter_count(self) -> int:
        return self.com_object.ParameterCount

    @property
    def parameter_types(self) -> list:
        return self.com_object.ParameterTypes

    def call(self, *parameters):
        return self.com_object.Call(*parameters)
