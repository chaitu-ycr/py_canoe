# Import Python Libraries here
import win32com.client
from typing import Union

class Bus:
    """The Bus object represents a bus of the CANoe application.
    """
    def __init__(self, app_obj, bus_type='CAN') -> None:
        """Returns a Signal object.

        Args:
            app_obj (object): application class instance object.
            bus_type (str, optional): The desired bus type. Valid types are: CAN, LIN, FlexRay, AFDX, Ethernet. Defaults to 'CAN'.
        """
        self.app_obj = app_obj
        self.log = self.app_obj.log
        self.bus_com_obj = self.app_obj.app_com_obj.GetBus(bus_type)
    
    @property
    def active(self) -> bool:
        """determines the status of the Bus object.

        Returns:
            bool: The status of the Bus object.
        """
        return self.bus_com_obj.Active
    
    @active.setter
    def active(self, value: bool) -> None:
        """Sets the status of the Bus object.

        Args:
            value (bool): A boolean value that indicates whether the bus is to be simulated: TRUE: The bus will be simulated. FALSE: The bus will not be simulated.
        """
        self.bus_com_obj.Active = value
        self.log.info(f'status of the Bus object set to {value}.')
    
    @property
    def baudrate(self, channel_number: int) -> int:
        """Determines the baud rate of a channel.

        Args:
            channel_number (int): The channel number.

        Returns:
            int: The current baud rate of the channel.
        """
        return self.bus_com_obj.Baudrate(channel_number)
    
    @property
    def bus_name(self) -> str:
        """returns the bus name.

        Returns:
            str: The bus name.
        """
        return self.bus_com_obj.Name
    
    def set_bus_name(self, name: str) -> None:
        """Sets the bus name

        Args:
            name (str): The new name.
        """
        self.bus_com_obj.Name = name
        self.log.info(f'bus name set to {name}.')
    
    def get_signal(self, channel: int, message: str, signal: str) -> object:
        """Returns a Signal object.

        Args:
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.

        Returns:
            object: The Signal object.
        """
        return self.bus_com_obj.GetSignal(channel, message, signal)
    
    def get_j1939_signal(self, channel: int, message: str, signal: str, source_address: int, destination_address: int) -> object:
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
        return self.bus_com_obj.GetJ1939Signal(channel, message, signal, source_address, destination_address)

class Channel:
    def __init__(self) -> None:
        pass

class Database:
    def __init__(self) -> None:
        pass

class Generator:
    def __init__(self) -> None:
        pass

class Node:
    def __init__(self) -> None:
        pass

class ReplayBlock:
    def __init__(self) -> None:
        pass

class Signal:
    """The Signal object represents a signal on the bus.
    """
    def __init__(self, signal_com_object) -> None:
        self.sig_com_obj = signal_com_object
    
    @property
    def full_name(self) -> str:
        """Determines the fully qualified name of a signal.

        Returns:
            str: The fully qualified name of a signal or a message. The following format will be used for signals: <DatabaseName>::<MessageName>::<SignalName>
        """
        return self.sig_com_obj.FullName

    @property
    def is_online(self) -> bool:
        """Checks whether the measurement is running and the signal has been received.

        Returns:
            bool: TRUE: if the measurement is running and the signal has been received. FALSE: if not.
        """
        return self.sig_com_obj.IsOnline

    @property
    def raw_value(self) -> int:
        """Returns the current value of the signal as it was transmitted on the bus.

        Returns:
            int: The raw value of the signal.
        """
        return self.sig_com_obj.RawValue
    
    @raw_value.setter
    def raw_value(self, value: int) -> None:
        """Returns the current value of the signal as it was transmitted on the bus.

        Returns:
            int: The raw value of the signal.
        """
        self.sig_com_obj.RawValue = value

    @property
    def state(self) -> int:
        """Returns the state of the signal.

        Returns:
            int: State of the signal; possible values are: 0: The default value of the signal is returned. 1: The measurement is not running; the value set by the application is returned. 3: The signal has been received in the current measurement; the current value is returned.
        """
        return self.sig_com_obj.State

    @property
    def value(self) -> int:
        """gets the active value of the signal.


        Returns:
            int: The value of the signal
        """
        return self.sig_com_obj.Value
    
    @value.setter
    def value(self, value: int) -> None:
        """sets the active value of the signal.

        Args:
            value (int): The new value of the signal.
        """
        self.sig_com_obj.Value = value
