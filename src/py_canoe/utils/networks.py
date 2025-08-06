import win32com.client
from typing import Union
from py_canoe.utils.common import logger
from py_canoe.utils.common import DoEventsUntil


class DiagnosticRequestEvents:
    TIMEOUT = False
    RECEIVED_RESPONSE = False
    RESPONSE: Union['DiagnosticResponse', None] = None

    @staticmethod
    def OnCompletion():
        DiagnosticRequestEvents.TIMEOUT = False
        DiagnosticRequestEvents.RECEIVED_RESPONSE = False
        DiagnosticRequestEvents.RESPONSE = None

    @staticmethod
    def OnConfirmation():
        DiagnosticRequestEvents.TIMEOUT = False
        DiagnosticRequestEvents.RECEIVED_RESPONSE = False
        DiagnosticRequestEvents.RESPONSE = None


    @staticmethod
    def OnResponse(response):
        DiagnosticRequestEvents.TIMEOUT = False
        DiagnosticRequestEvents.RECEIVED_RESPONSE = True
        DiagnosticRequestEvents.RESPONSE = DiagnosticResponse(response)

    @staticmethod
    def OnTimeout():
        DiagnosticRequestEvents.TIMEOUT = True
        DiagnosticRequestEvents.RECEIVED_RESPONSE = False
        DiagnosticRequestEvents.RESPONSE = None


class Diagnostic:
    def __init__(self, diagnostic):
        self.com_object = diagnostic

    @property
    def tester_present_status(self) -> bool:
        return self.com_object.TesterPresentStatus

    def create_request(self, primitive_path) -> 'DiagnosticRequest':
        return DiagnosticRequest(self.com_object.CreateRequest(primitive_path))

    def create_request_from_stream(self, byte_stream: bytearray) -> 'DiagnosticRequest':
        return DiagnosticRequest(self.com_object.CreateRequestFromStream(byte_stream))

    def diag_start_tester_present(self):
        self.com_object.DiagStartTesterPresent()

    def diag_stop_tester_present(self):
        self.com_object.DiagStopTesterPresent()


class DiagnosticRequest:
    def __init__(self, diagnostic_request, enable_events: bool = True):
        self.com_object = diagnostic_request
        DiagnosticRequestEvents.TIMEOUT = False
        DiagnosticRequestEvents.RECEIVED_RESPONSE = False
        self.wait_for_response_or_timeout = lambda: DoEventsUntil(lambda: (DiagnosticRequestEvents.RECEIVED_RESPONSE or DiagnosticRequestEvents.TIMEOUT), lambda: 300, "Diagnostic Request Response")
        if enable_events:
            win32com.client.WithEvents(self.com_object, DiagnosticRequestEvents)

    @property
    def pending(self) -> bool:
        return self.com_object.Pending

    @property
    def responses(self) -> 'DiagnosticResponses':
        return DiagnosticResponses(self.com_object.Responses)

    @property
    def suppress_positive_response(self) -> bool:
        return self.com_object.SuppressPositiveResponse

    @suppress_positive_response.setter
    def suppress_positive_response(self, value: bool):
        self.com_object.SuppressPositiveResponse = value

    def send(self):
        self.com_object.Send()
        self.wait_for_response_or_timeout()

    def set_complex_parameter(self, qualifier, iteration, sub_parameter, value):
        self.com_object.SetComplexParameter(qualifier, iteration, sub_parameter, value)

    def set_parameter(self, qualifier, value):
        self.com_object.SetParameter(qualifier, value)


class DiagnosticResponses:
    def __init__(self, diagnostic_responses):
        self.com_object = win32com.client.Dispatch(diagnostic_responses)

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'DiagnosticResponse':
        return DiagnosticResponse(self.com_object.item(index))


class DiagnosticResponse:
    def __init__(self, diagnostic_response):
        self.com_object = win32com.client.Dispatch(diagnostic_response)

    @property
    def positive(self) -> bool:
        return self.com_object.Positive

    @property
    def response_code(self) -> int:
        return self.com_object.ResponseCode

    @property
    def sender(self) -> str:
        return self.com_object.Sender

    @property
    def stream(self) -> bytearray:
        return self.com_object.Stream

    def get_complex_iteration_count(self, qualifier: str) -> int:
        return self.com_object.GetComplexIterationCount(qualifier)

    def get_complex_parameter(self, qualifier: str, iteration: int, sub_parameter: str, mode: int) -> Union[bytearray, int, str]:
        return self.com_object.GetComplexParameter(qualifier, iteration, sub_parameter, mode)

    def get_parameter(self, qualifier: str, mode: int) -> Union[bytearray, int, str]:
        return self.com_object.GetParameter(qualifier, mode)

    def is_complex_parameter(self, qualifier: str) -> bool:
        return self.com_object.IsComplexParameter(qualifier)
