# import external modules here
import logging
import pythoncom
import win32com.client


class Networks:
    """The Networks class represents the networks of CANoe."""
    def __init__(self, app_com_obj):
        try:
            self.log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.Networks)
        except Exception as e:
            self.log.error(f'😡 Error initializing Networks object: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def fetch_all_networks(self) -> dict:
        networks = dict()
        for index in range(1, self.count + 1):
            network_com_obj = win32com.client.Dispatch(self.com_obj.Item(index))
            network = Network(network_com_obj)
            networks[network_com_obj.Name] = network
        return networks

    def fetch_all_diag_devices(self) -> dict:
        diag_devices = dict()
        networks = self.fetch_all_networks()
        if len(networks) > 0:
            for _, n_value in networks.items():
                devices = n_value.devices
                n_devices = devices.get_all_devices()
                if len(n_devices) > 0:
                    for d_name, d_value in n_devices.items():
                        if d_value.diagnostic is not None:
                            diag_devices[d_name] = d_value.diagnostic
        return diag_devices


class Network:
    """The Network class represents one single network of CANoe."""
    def __init__(self, network_com_obj):
        self.com_obj = network_com_obj

    @property
    def bus_type(self) -> int:
        return self.com_obj.BusType

    @property
    def devices(self) -> object:
        return Devices(self.com_obj)

    @property
    def name(self) -> str:
        return self.com_obj.Name


class Devices:
    """The Devices class represents all devices of CANoe."""
    def __init__(self, network_com_obj):
        self.com_obj = network_com_obj.Devices

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def get_all_devices(self) -> dict:
        devices = dict()
        for index in range(1, self.count + 1):
            device_com_obj = self.com_obj.Item(index)
            device = Device(device_com_obj)
            devices[device.name] = device
        return devices


class Device:
    """The Device class represents one single device of CANoe."""
    def __init__(self, device_com_obj):
        self.com_obj = device_com_obj

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def diagnostic(self):
        try:
            diag_com_obj = self.com_obj.Diagnostic
            return Diagnostic(diag_com_obj)
        except pythoncom.com_error:
            return None


class Diagnostic:
    """The Diagnostic class represents the diagnostic properties of an ECU on the bus or the basic diagnostic functionality of a CANoe network.
    It is identified by the ECU qualifier that has been specified for the loaded diagnostic description (CDD/ODX).
    """
    def __init__(self, diagnostic_com_obj):
        self.com_obj = diagnostic_com_obj

    @property
    def tester_present_status(self) -> bool:
        return self.com_obj.TesterPresentStatus

    def create_request(self, primitive_path: str):
        return DiagnosticRequest(self.com_obj.CreateRequest(primitive_path))

    def create_request_from_stream(self, byte_stream: str):
        diag_req_in_bytes = bytearray()
        byte_stream = ''.join(byte_stream.split(' '))
        for i in range(0, len(byte_stream), 2):
            diag_req_in_bytes.append(int(byte_stream[i:i + 2], 16))
        return DiagnosticRequest(self.com_obj.CreateRequestFromStream(diag_req_in_bytes))

    def start_tester_present(self) -> None:
        self.com_obj.DiagStartTesterPresent()

    def stop_tester_present(self) -> None:
        self.com_obj.DiagStopTesterPresent()


class DiagnosticRequest:
    """The DiagnosticRequest class represents the query of a diagnostic tester (client) to an ECU (server) in CANoe.
    It can be replied by a DiagnosticResponse object.
    """
    def __init__(self, diag_req_com_obj):
        self.com_obj = diag_req_com_obj

    @property
    def pending(self) -> bool:
        return self.com_obj.Pending

    @property
    def responses(self) -> list:
        diag_responses_com_obj = self.com_obj.Responses
        diag_responses = [DiagnosticResponse(diag_responses_com_obj.item(i)) for i in range(1, diag_responses_com_obj.Count + 1)]
        return diag_responses

    @property
    def suppress_positive_response(self):
        return self.com_obj.SuppressPositiveResponse

    def send(self):
        self.com_obj.Send()

    def set_complex_parameter(self, qualifier, iteration, sub_parameter, value):
        self.com_obj.SetComplexParameter(qualifier, iteration, sub_parameter, value)

    def set_parameter(self, qualifier, value):
        self.com_obj.SetParameter(qualifier, value)


class DiagnosticResponse:
    """The DiagnosticResponse class represents the ECU's reply to a diagnostic request in CANoe.
    The received parameters can be read out and processed.
    """
    def __init__(self, diag_res_com_obj):
        self.com_obj = diag_res_com_obj

    @property
    def positive(self) -> bool:
        return self.com_obj.Positive

    @property
    def response_code(self) -> int:
        return self.com_obj.ResponseCode

    @property
    def stream(self) -> str:
        diag_response_data = " ".join(f"{d:02X}" for d in self.com_obj.Stream).upper()
        return diag_response_data

    @property
    def sender(self) -> str:
        return self.com_obj.Sender

    def get_complex_iteration_count(self, qualifier):
        return self.com_obj.GetComplexIterationCount(qualifier)

    def get_complex_parameter(self, qualifier, iteration, sub_parameter, mode):
        return self.com_obj.GetComplexParameter(qualifier, iteration, sub_parameter, mode)

    def get_parameter(self, qualifier, mode):
        return self.com_obj.GetParameter(qualifier, mode)

    def is_complex_parameter(self, qualifier):
        return self.com_obj.IsComplexParameter(qualifier)
