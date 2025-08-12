import win32com.client
from typing import Union
from py_canoe.utils.common import logger, wait
from py_canoe.utils.common import DoEventsUntil

DIAGNOSTIC_RESPONSE_TIMEOUT_VALUE = 300 # 5 minutes


class DiagnosticResponse:
    def __init__(self, diagnostic_response):
        self.com_object = diagnostic_response

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
        return bytearray(self.com_object.Stream)

    def get_complex_iteration_count(self, qualifier: str) -> int:
        return self.com_object.GetComplexIterationCount(qualifier)

    def get_complex_parameter(self, qualifier: str, iteration: int, sub_parameter: str, mode: int) -> Union[bytearray, int, str]:
        return self.com_object.GetComplexParameter(qualifier, iteration, sub_parameter, mode)

    def get_parameter(self, qualifier: str, mode: int) -> Union[bytearray, int, str]:
        return self.com_object.GetParameter(qualifier, mode)

    def is_complex_parameter(self, qualifier: str) -> bool:
        return self.com_object.IsComplexParameter(qualifier)


class DiagnosticResponses:
    def __init__(self, diagnostic_responses):
        self.com_object = diagnostic_responses

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> DiagnosticResponse:
        return DiagnosticResponse(self.com_object.item(index))


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


class DiagnosticRequest:
    def __init__(self, diagnostic_request, enable_events: bool = True):
        self.com_object = diagnostic_request
        if enable_events:
            win32com.client.WithEvents(self.com_object, DiagnosticRequestEvents)

    @property
    def pending(self) -> bool:
        return self.com_object.Pending

    @property
    def responses(self) -> DiagnosticResponses:
        return DiagnosticResponses(self.com_object.Responses)

    @property
    def suppress_positive_response(self) -> bool:
        return self.com_object.SuppressPositiveResponse

    @suppress_positive_response.setter
    def suppress_positive_response(self, value: bool):
        self.com_object.SuppressPositiveResponse = value

    @staticmethod
    def _condition():
        return DiagnosticRequestEvents.RECEIVED_RESPONSE or DiagnosticRequestEvents.TIMEOUT

    def _wait_for_response_or_timeout(self) -> bool:
        DiagnosticRequestEvents.TIMEOUT = False
        DiagnosticRequestEvents.RECEIVED_RESPONSE = False
        status = DoEventsUntil(self._condition, DIAGNOSTIC_RESPONSE_TIMEOUT_VALUE, "Diagnostic Request Response")
        if not status:
            logger.error(f"😡 Error: Diagnostic request did not receive a response within {DIAGNOSTIC_RESPONSE_TIMEOUT_VALUE} seconds.")
        return status

    def send(self):
        self.com_object.Send()
        self._wait_for_response_or_timeout()

    def set_complex_parameter(self, qualifier, iteration, sub_parameter, value):
        self.com_object.SetComplexParameter(qualifier, iteration, sub_parameter, value)

    def set_parameter(self, qualifier, value):
        self.com_object.SetParameter(qualifier, value)


class Diagnostic:
    def __init__(self, diagnostic):
        self.com_object = diagnostic

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


class ApplicationSocket:
    def __init__(self, application_socket):
        self.com_object = application_socket

    @property
    def bus_registry(self) -> bytearray:
        return self.com_object.BusRegistry

    @property
    def fb_lock_list(self) -> bytearray:
        return self.com_object.FBlockList


class AudioInterface:
    def __init__(self, audio_interface):
        self.com_object = audio_interface

    def mute(self, line_in_out: int, mute: Union[int, None]=None) -> int:
        if mute is None:
            return self.com_object.Mute(line_in_out)
        else:
            obj = self.com_object.Mute(line_in_out)
            obj = mute
            return self.com_object.Mute(line_in_out)

    def volume(self, line_in_out: int, volume: Union[int, None]=None) -> int:
        if volume is None:
            return self.com_object.Volume(line_in_out)
        else:
            obj = self.com_object.Volume(line_in_out)
            obj = volume
            return self.com_object.Volume(line_in_out)

    def connect_to_label(self, line_in_out: int, connection_label: int):
        self.com_object.ConnectToLabel(line_in_out, connection_label)

    def disconnect_from_label(self, line_in_out: int, connection_label: int):
        self.com_object.DisconnectFromLabel(line_in_out, connection_label)


class MostDisassembler:
    def __init__(self, disassembler):
        self.com_object = disassembler

    def symbolic_message_id_components(self, f_block_id: int, instance_id: int, function_id: int, op_type_id: int) -> int:
        return self.com_object.SymbolicMessageIDComponents(f_block_id, instance_id, function_id, op_type_id)

    def symbolic_parameter_list1(self, data_length: int, data_array: bytearray, max_params: int = 0) -> tuple:
        return self.com_object.SymbolicParameterList1(data_length, data_array, max_params)

    def symbolic_parameter_list2(self, f_block_id: int, instance_id: int, function_id: int, op_type_id: int, data_length: int, data_array: bytearray, max_params: int = 0) -> tuple:
        return self.com_object.SymbolicParameterList2(f_block_id, instance_id, function_id, op_type_id, data_length, data_array, max_params)

    def this_message_id_components(self, f_block_id: int, instance_id: int, function_id: int, op_type_id: int) -> int:
        return self.com_object.ThisMessageIDComponents(f_block_id, instance_id, function_id, op_type_id)

    def this_symbolic_message_id_components(self, f_block_name: str, function_name: str, op_type_name: str) -> int:
        return self.com_object.ThisSymbolicMessageIDComponents(f_block_name, function_name, op_type_name)

    def this_symbolic_parameter_list(self, max_params: int = 0) -> tuple:
        return self.com_object.ThisSymbolicParameterList(max_params)


class MostNetworkInterface:
    def __init__(self, network_interface):
        # TODO: Implement the MostNetworkInterface class later if required
        self.com_object = network_interface


class Device:
    def __init__(self, device):
        self.com_object = device

    @property
    def application_socket(self) -> ApplicationSocket:
        return ApplicationSocket(self.com_object.ApplicationSocket)

    @property
    def audio_interface(self) -> AudioInterface:
        return AudioInterface(self.com_object.AudioInterface)

    @property
    def diagnostic(self) -> Diagnostic:
        return Diagnostic(self.com_object.Diagnostic)

    @property
    def disassembler(self) -> MostDisassembler:
        return MostDisassembler(self.com_object.Disassembler)

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def network_interface(self) -> MostNetworkInterface:
        return MostNetworkInterface(self.com_object.NetworkInterface)


class Devices:
    def __init__(self, devices):
        self.com_object = devices

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'Device':
        return Device(self.com_object.Item(index))


class Network:
    def __init__(self, network):
        self.com_object = network

    @property
    def bus_type(self) -> int:
        return self.com_object.BusType

    @property
    def devices(self) -> Devices:
        return Devices(self.com_object.Devices)

    @property
    def name(self) -> str:
        return self.com_object.Name


class Networks:
    """
    The Networks object represents the networks of CANoe.
    """
    def __init__(self, app):
        self.com_object = app.com_object.Networks

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> Network:
        return Network(self.com_object.Item(index))


def fetch_diagnostic_devices(app):
    try:
        app._diagnostic_devices = {}
        for i in range(1, app.com_object.Networks.Count + 1):
            network = app.com_object.Networks.Item(i)
            for j in range(1, network.Devices.Count + 1):
                device = network.Devices.Item(j)
                try:
                    diagnostic = getattr(device, 'Diagnostic', None)
                    if diagnostic:
                        app._diagnostic_devices[device.Name] = Diagnostic(diagnostic)
                except Exception:
                    pass
    except Exception as e:
        logger.error(f"😡 Error fetching Diagnostic Devices: {e}")
        return None

def send_diag_request(app, diag_ecu_qualifier_name: str, request: str, request_in_bytes=True, return_sender_name=False, response_in_bytearray=False) -> Union[str, dict]:
    try:
        diag_device: Diagnostic = app._diagnostic_devices.get(diag_ecu_qualifier_name)
        if diag_device:
            if request_in_bytes:
                diag_req_in_bytes = bytearray()
                byte_stream = ''.join(request.split(' '))
                for i in range(0, len(byte_stream), 2):
                    diag_req_in_bytes.append(int(byte_stream[i:i + 2], 16))
                diag_request = diag_device.create_request_from_stream(diag_req_in_bytes)
            else:
                diag_request = diag_device.create_request(request)
            diag_request.send()
            logger.info(f'💉 {diag_ecu_qualifier_name}: Diagnostic Request = {request}')
            while diag_request.pending:
                wait(0.01)
            diag_responses_dict = {}
            diag_response_including_sender_name = {}
            for i in range(1, diag_request.responses.count + 1):
                diag_response = diag_request.responses.item(i)
                diag_response_positive = diag_response.positive
                response_code = diag_response.response_code
                response_sender = diag_response.sender
                response_stream = diag_response.stream
                response_stream_in_str = " ".join(f"{d:02X}" for d in response_stream).upper()
                diag_responses_dict[response_sender] = {
                    "positive": diag_response_positive,
                    "response_code": response_code,
                    "stream": response_stream,
                    "stream_in_str": response_stream_in_str
                }
                if response_in_bytearray:
                    diag_response_including_sender_name[response_sender] = response_stream
                else:
                    diag_response_including_sender_name[response_sender] = response_stream_in_str
                if diag_response_positive:
                    logger.info(f'🟢 {response_sender}: Diagnostic Response = {response_stream_in_str}')
                else:
                    logger.info(f'🔴 {response_sender}: Diagnostic Response = {response_stream_in_str}')
            return diag_response_including_sender_name if return_sender_name else diag_response_including_sender_name[diag_ecu_qualifier_name]
        else:
            logger.warning(f'⚠️ No responses received for request: {request}')
            return {"error": "No responses received"}
    except Exception as e:
        logger.error(f"😡 Error sending diagnostic request: {e}")
        return {"error": str(e)}

def control_tester_present(app, diag_ecu_qualifier_name: str, value: bool) -> bool:
    try:
        diag_device: Diagnostic = app._diagnostic_devices.get(diag_ecu_qualifier_name)
        if diag_device:
            if value:
                diag_device.diag_start_tester_present()
                logger.info(f'✔️ {diag_ecu_qualifier_name}: Tester Present started.')
            else:
                diag_device.diag_stop_tester_present()
                logger.info(f'⏹️ {diag_ecu_qualifier_name}: Tester Present stopped.')
            return True
        else:
            logger.warning(f'⚠️ No diagnostic device found for: {diag_ecu_qualifier_name}')
            return False
    except Exception as e:
        logger.error(f"😡 Error controlling tester present: {e}")
        return False
