# import external modules here
import logging
import pythoncom
import win32com.client

# import internal modules here


class Networks:
    """The Networks class represents the networks of CANoe.
    """
    def __init__(self, app_com_obj):
        try:
            self.log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.Networks)
        except Exception as e:
            self.log.error(f'😡 Error initializing Networks object: {str(e)}')

    @property
    def count(self) -> int:
        """Returns the number of Networks inside the collection.

        Returns:
            int: The number of Networks contained
        """
        return self.com_obj.Count

    def fetch_all_networks(self) -> dict:
        """returns all networks available in configuration.
        """
        networks = dict()
        for index in range(1, self.count + 1):
            network_com_obj = win32com.client.Dispatch(self.com_obj.Item(index))
            network = Network(network_com_obj)
            networks[network_com_obj.Name] = network
        return networks

    def fetch_all_diag_devices(self) -> dict:
        """returns all diagnostic devices available in configuration.
        """
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
    """The Network class represents one single network of CANoe.
    """
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
        """Returns the Devices class.

        Returns:
            object: The Devices object
        """
        return Devices(self.com_obj)

    @property
    def name(self) -> str:
        """The name of the network.
        """
        return self.com_obj.Name


class Devices:
    """The Devices class represents all devices of CANoe.
    """
    def __init__(self, network_com_obj):
        self.com_obj = network_com_obj.Devices

    @property
    def count(self) -> int:
        """Returns the number of Networks inside the collection.

        Returns:
            int: The number of Networks contained
        """
        return self.com_obj.Count

    def get_all_devices(self) -> dict:
        devices = dict()
        for index in range(1, self.count + 1):
            device_com_obj = self.com_obj.Item(index)
            device = Device(device_com_obj)
            devices[device.name] = device
        return devices


class Device:
    """The Device class represents one single device of CANoe.
    """
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
        """The Diagnostic object represents the diagnostic properties of an ECU on the bus or the basic diagnostic functionality of a CANoe network.
        It is identified by the ECU qualifier that has been specified for the loaded diagnostic description (CDD/ODX).
        """
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
        return DiagnosticRequest(self.com_obj.CreateRequest(primitive_path))

    def create_request_from_stream(self, byte_stream: str):
        """Creates a request object with the given byte stream.
        If no request for the given byte sequence can be found in the diagnostic description (CDD/ODX), it is not possible to access the parameters symbolically.
        """
        diag_req_in_bytes = bytearray()
        byte_stream = ''.join(byte_stream.split(' '))
        for i in range(0, len(byte_stream), 2):
            diag_req_in_bytes.append(int(byte_stream[i:i + 2], 16))
        return DiagnosticRequest(self.com_obj.CreateRequestFromStream(diag_req_in_bytes))

    def start_tester_present(self) -> None:
        """Starts sending autonomous/cyclical Tester Present requests to the ECU.
        The TesterPresent remains active only as long as the COM script is running.
        When the COM script is finished, the diagnostic channel is closed and the Tester Present process is switched off.
        """
        self.com_obj.DiagStartTesterPresent()

    def stop_tester_present(self) -> None:
        """stops sending autonomous/cyclical Tester Present requests to the ECU.
        """
        self.com_obj.DiagStopTesterPresent()


class DiagnosticRequest:
    """The DiagnosticRequest class represents the query of a diagnostic tester (client) to an ECU (server) in CANoe.
    It can be replied by a DiagnosticResponse object.
    """
    def __init__(self, diag_req_com_obj):
        self.com_obj = diag_req_com_obj

    @property
    def pending(self) -> bool:
        """The Pending state of a request is True, as long as events may occur for it. For this a request must have been sent.
        read-only. The initial value of this property is False.
        """
        return self.com_obj.Pending

    @property
    def responses(self) -> list:
        """Contains the diagnostic response objects received for the request object.
        If the request has not been sent yet, the container is empty.
        """
        diag_responses_com_obj = self.com_obj.Responses
        diag_responses = [DiagnosticResponse(diag_responses_com_obj.item(i)) for i in range(1, diag_responses_com_obj.Count + 1)]
        return diag_responses

    @property
    def suppress_positive_response(self):
        """A boolean value.
        The ECU will return a negative response only in case of an error.
        """
        return self.com_obj.SuppressPositiveResponse

    def send(self):
        """Causes the DiagnosticRequest object to be sent.
        """
        self.com_obj.Send()

    def set_complex_parameter(self, qualifier, iteration, sub_parameter, value):
        """Sets the value of the parameter of the request to the given value.
        """
        self.com_obj.SetComplexParameter(qualifier, iteration, sub_parameter, value)

    def set_parameter(self, qualifier, value):
        """Sets the value of the parameter of the request to the given value.
        """
        self.com_obj.SetParameter(qualifier, value)


class DiagnosticResponse:
    """The DiagnosticResponse class represents the ECU's reply to a diagnostic request in CANoe.
    The received parameters can be read out and processed.
    """
    def __init__(self, diag_res_com_obj):
        self.com_obj = diag_res_com_obj

    @property
    def positive(self) -> bool:
        """True for positive responses.
        False for negative responses.
        """
        return self.com_obj.Positive

    @property
    def response_code(self) -> int:
        """The error code (integer) provided by the ECU.
        """
        return self.com_obj.ResponseCode

    @property
    def stream(self) -> str:
        """The byte sequence of the object as received.
        """
        diag_response_data = " ".join(f"{d:02X}" for d in self.com_obj.Stream).upper()
        return diag_response_data

    @property
    def sender(self) -> str:
        """The identifier (string) of the ECU that sent the response.
        """
        return self.com_obj.Sender

    def get_complex_iteration_count(self, qualifier):
        """The number of iterations the parameter holds, e.g. the number of DTCs in a list.
        For simple parameters, 0 is returned.
        For a complex parameter that is no iteration, e.g. a structure, 1 is returned.
        """
        return self.com_obj.GetComplexIterationCount(qualifier)

    def get_complex_parameter(self, qualifier, iteration, sub_parameter, mode):
        """Value of the parameter as a string (default), as a raw value or as a byte sequence.
        """
        return self.com_obj.GetComplexParameter(qualifier, iteration, sub_parameter, mode)

    def get_parameter(self, qualifier, mode):
        """Value of the parameter as a string (default), as a raw value or as a byte sequence.
        """
        return self.com_obj.GetParameter(qualifier, mode)

    def is_complex_parameter(self, qualifier):
        """TRUE if parameter has sub-parameters.
        FALSE if parameter has no sub-parameters.
        """
        return self.com_obj.IsComplexParameter(qualifier)
