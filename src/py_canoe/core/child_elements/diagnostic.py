from py_canoe.core.child_elements.diagnostic_request import DiagnosticRequest


class Diagnostic:
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def tester_present_status(self) -> bool:
        return self.com_object.TesterPresentStatus

    def create_request(self, primitive_path) -> DiagnosticRequest:
        return DiagnosticRequest(self.com_object.CreateRequest(primitive_path))

    def create_request_from_stream(self, byte_stream: bytearray) -> DiagnosticRequest:
        return DiagnosticRequest(self.com_object.CreateRequestFromStream(byte_stream))

    def diag_start_tester_present(self):
        self.com_object.DiagStartTesterPresent()

    def diag_stop_tester_present(self):
        self.com_object.DiagStopTesterPresent()
