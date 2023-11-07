# Import Python Libraries here
import logging
import pythoncom
import win32com.client


class Networks:
    def __init__(self, app_com_obj):
        self.log = logging.getLogger('CANOE_LOG')
        self.com_obj = win32com.client.Dispatch(app_com_obj.Networks)

    @property
    def count(self) -> int:
        """Returns the number of Networks inside the collection.

        Returns:
            int: The number of Networks contained
        """
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
            for n_name, n_value in networks.items():
                devices = n_value.devices
        return diag_devices

    def fetch_diag_devices(self) -> dict:
        diag_devices = dict()
        for network in self.com_obj:
            for device in network.Devices:
                try:
                    diag_devices[device.Name] = device.Diagnostic
                except pythoncom.com_error:
                    pass
        return diag_devices


class Network:
    def __init__(self, network_com_obj):
        self.com_obj = network_com_obj

    @property
    def bus_type(self) -> int:
        """Determines the bus type of the network

        Returns:
            int: The type of the network: 0-Invalid, 1-CAN, 2-J1939, 5-LIN, 6-MOST, 7-FlexRay, 9-J1708, 11-Ethernet
        """
        return self.com_obj.BusType

    @property
    def devices(self) -> object:
        """Returns the Devices object.

        Returns:
            object: The Devices object
        """
        return Devices(self.com_obj)

    @property
    def name(self) -> str:
        return self.com_obj.Name


class Devices:
    def __init__(self, network_com_obj):
        self.com_obj = network_com_obj.Devices

    @property
    def count(self) -> int:
        """Returns the number of Networks inside the collection.

        Returns:
            int: The number of Networks contained
        """
        return self.com_obj.Count

    def get_all_devices(self):
        devices = dict()
        for index in range(1, self.count + 1):
            device_com_obj = self.com_obj.Item(index)
            device = Device(device_com_obj)
            devices[device_com_obj.Name] = device
        return devices


class Device:
    def __init__(self, device_com_obj):
        self.com_obj = device_com_obj

    @property
    def name(self) -> str:
        """The name of the device.

        Returns:
            str: The name of the device.
        """
        return self.com_obj.Name

    @property
    def diagnostic(self):
        try:
            diag_com_obj = self.com_obj.Diagnostic
            return Diagnostic(diag_com_obj)
        except pythoncom.com_error:
            return None


class Diagnostic:
    def __init__(self, diagnostic_com_obj):
        self.com_obj = diagnostic_com_obj

    @property
    def tester_present_status(self) -> bool:
        """Returns the status of autonomous/cyclical Tester Present requests to the ECU.
        The status of autonomous/cyclical Tester Present requests to the ECU:
        true: Sending Tester Present enabled for this ECU.
        false: Sending Tester Present disabled for this ECU.
        """
        return self.com_obj.TesterPresentStatus

    def create_request(self, primitive_path: str):
        """Creates a request object with given qualifier path.
        It is not possible to create response objects since they can only be generated by a responding ECU!
        """
        return self.com_obj.CreateRequest(primitive_path)

    def create_request_from_stream(self, byte_stream: str):
        """Creates a request object with the given byte stream.
        If no request for the given byte sequence can be found in the diagnostic description (CDD/ODX), it is not possible to access the parameters symbolically.
        """
        return self.com_obj.CreateRequestFromStream(byte_stream)

    def diag_start_tester_present(self) -> None:
        """Starts sending autonomous/cyclical Tester Present requests to the ECU.
        The TesterPresent remains active only as long as the COM script is running.
        When the COM script is finished, the diagnostic channel is closed and the Tester Present process is switched off.
        """
        self.com_obj.DiagStartTesterPresent()

    def diag_stop_tester_present(self) -> None:
        """stops sending autonomous/cyclical Tester Present requests to the ECU.
        """
        self.com_obj.DiagStopTesterPresent()


class DiagnosticRequest:
    def __init__(self, diag_req_com_obj):
        self.com_obj = diag_req_com_obj

    @property
    def pending(self) -> bool:
        """The Pending state of a request is True, as long as events may occur for it. For this a request must have been sent.
        read-only. The initial value of this property is False.
        """
        return self.com_obj.Pending

    @property
    def responses(self):
        return self.com_obj.Responses

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
    def __init__(self, diag_res_com_obj):
        self.com_obj = diag_res_com_obj

    @property
    def positive(self) -> bool:
        return self.com_obj.Positive

    @property
    def response_code(self) -> int:
        return self.com_obj.ResponseCode

    @property
    def stream(self) -> tuple:
        return self.com_obj.Stream

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
