# Import Python Libraries here
import logging


class Bus:
    """The Bus object represents a bus of the CANoe application.
    """

    def __init__(self, app_com_obj, bus_type='CAN'):
        """Returns a Signal object.

        Args:
            app_com_obj (object): application com object.
            bus_type (str, optional): The desired bus type. Valid types are: CAN, LIN, FlexRay, AFDX, Ethernet. Defaults to 'CAN'.
        """
        self.app_com_obj = app_com_obj
        self.log = logging.getLogger('CANOE_LOG')
        self.com_obj = self.app_com_obj.GetBus(bus_type)

    def reinit_bus(self, bus_type='CAN'):
        self.com_obj = self.app_com_obj.GetBus(bus_type)

    def get_signal(self, channel: int, message: str, signal: str) -> object:
        """Returns a Signal object.

        Args:
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.

        Returns:
            object: The Signal object.
        """
        return self.com_obj.GetSignal(channel, message, signal)

    def get_j1939_signal(self, channel: int, message: str, signal: str, source_address: int,
                         destination_address: int) -> object:
        """Returns a Signal object.

        Args:
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.
            source_address (int): The source address of the ECU that sends the message
            destination_address (int): The destination address of the ECU that receives the message. For signals of global Parameter Groups (PDU 2 format) this parameter is not considered.

        Returns:
            object: The Signal object.
        """
        return self.com_obj.GetJ1939Signal(channel, message, signal, source_address, destination_address)

    # Signal object relevant
    @staticmethod
    def signal_full_name(signal_object) -> str:
        """Determines the fully qualified name of a signal.

        Args:
            signal_object (object): signal object.

        Returns:
            str: The fully qualified name of a signal or a message. The following format will be used for signals: <DatabaseName>::<MessageName>::<SignalName>
        """
        return signal_object.FullName

    @staticmethod
    def signal_is_online(signal_object) -> bool:
        """Checks whether the measurement is running and the signal has been received.

        Args:
            signal_object (object): signal object.

        Returns:
            bool: TRUE: if the measurement is running and the signal has been received. FALSE: if not.
        """
        return signal_object.IsOnline

    @staticmethod
    def signal_get_raw_value(signal_object) -> int:
        """Returns the current value of the signal as it was transmitted on the bus.

        Args:
            signal_object (object): signal object.

        Returns:
            int: The raw value of the signal.
        """
        return signal_object.RawValue

    @staticmethod
    def signal_set_raw_value(signal_object, value: int) -> None:
        """Returns the current value of the signal as it was transmitted on the bus.

        Args:
            signal_object (object): signal object.
            value (int): The new raw value of the signal.

        Returns:
            int: The raw value of the signal.
        """
        signal_object.RawValue = value

    @staticmethod
    def signal_state(signal_object) -> int:
        """Returns the state of the signal.

        Args:
            signal_object (object): signal object.

        Returns:
            int: State of the signal; possible values are: 0: The default value of the signal is returned. 1: The measurement is not running; the value set by the application is returned. 3: The signal has been received in the current measurement; the current value is returned.
        """
        return signal_object.State

    @staticmethod
    def signal_get_value(signal_object) -> int:
        """gets the active value of the signal.

        Args:
            signal_object (object): signal object.

        Returns:
            int: The value of the signal
        """
        return signal_object.Value

    @staticmethod
    def signal_set_value(signal_object, value: int) -> None:
        """sets the active value of the signal.

        Args:
            signal_object (object): signal object.
            value (int): The new value of the signal.
        """
        signal_object.Value = value

    # Databases
    def database_objects(self):
        db_objects = dict()
        databases = self.com_obj.Databases
        for index in range(1, databases.Count + 1):
            database = databases.Item(index)
            db_objects[database.Name] = database
        return db_objects
    
    # Nodes
    def node_objects(self):
        n_objects = dict()
        nodes = self.com_obj.Nodes
        for index in range(1, nodes.Count + 1):
            node = nodes.Item(index)
            n_objects[node.Name] = node
        return n_objects
