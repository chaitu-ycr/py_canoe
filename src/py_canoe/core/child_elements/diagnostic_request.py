from typing import Union

from py_canoe.core.child_elements.diagnostic_responses import DiagnosticResponses
from py_canoe.core.child_elements.diagnostic_response import DiagnosticResponse

DIAGNOSTIC_RESPONSE_TIMEOUT_VALUE = 300 # 5 minutes



class DiagnosticRequestEvents:
    def __init__(self) -> None:
        self.TIMEOUT: bool = False
        self.RECEIVED_RESPONSE: bool = False
        self.RESPONSE: Union['DiagnosticResponse', None] = None

    def OnCompletion(self):
        pass

    def OnConfirmation(self):
        pass

    def OnResponse(self, response):
        self.RECEIVED_RESPONSE = True
        self.RESPONSE = DiagnosticResponse(response)

    def OnTimeout(self):
        self.TIMEOUT = True


class DiagnosticRequest:
    def __init__(self, com_object):
        self.com_object = com_object
        # self.diagnostic_request_events: DiagnosticRequestEvents = win32com.client.WithEvents(self.com_object, DiagnosticRequestEvents)

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
        # DoEventsUntil(lambda: self.pending, DIAGNOSTIC_RESPONSE_TIMEOUT_VALUE, "Diagnostic Request Response")

    def set_complex_parameter(self, qualifier, iteration, sub_parameter, value):
        self.com_object.SetComplexParameter(qualifier, iteration, sub_parameter, value)

    def set_parameter(self, qualifier, value):
        self.com_object.SetParameter(qualifier, value)
