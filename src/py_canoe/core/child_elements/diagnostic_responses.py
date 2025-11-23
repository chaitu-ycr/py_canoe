from py_canoe.core.child_elements.diagnostic_response import DiagnosticResponse


class DiagnosticResponses:
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'DiagnosticResponse':
        return DiagnosticResponse(self.com_object.item(index))
