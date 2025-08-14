from typing import Union

from py_canoe.utils.common import logger


class Signal:
    def __init__(self, bus, channel: int, message: str, signal: str, source_address: int = None, destination_address: int = None):
        if source_address and destination_address:
            self.com_object = bus.com_object.GetJ1939Signal(channel, message, signal, source_address, destination_address)
        else:
            self.com_object = bus.com_object.GetSignal(channel, message, signal)

    @property
    def full_name(self) -> str:
        return self.com_object.FullName

    @property
    def is_online(self) -> bool:
        return self.com_object.IsOnline

    @property
    def raw_value(self) -> int:
        return self.com_object.RawValue

    @raw_value.setter
    def raw_value(self, value: int):
        self.com_object.RawValue = value

    @property
    def state(self) -> int:
        return self.com_object.State

    @property
    def value(self) -> Union[int, float]:
        return self.com_object.Value

    @value.setter
    def value(self, value: Union[int, float]):
        self.com_object.Value = value


class Bus:
    """
    The Bus object represents a bus of the CANoe application.
    """
    def __init__(self, app):
        self.app = app
        self.com_object = self.set_bus('CAN')
        self.VALUE_TABLE_SIGNAL_IS_ONLINE = {
            True: "measurement is running and the signal has been received.",
            False: "The signal is not online."
        }
        self.VALUE_TABLE_SIGNAL_STATE = {
            0: "The default value of the signal is returned.",
            1: "The measurement is not running. The value set by the application is returned.",
            2: "The measurement is not running. The value of the last measurement is returned.",
            3: "The signal has been received in the current measurement. The current value is returned."
        }

    def set_bus(self, bus_type: str = 'CAN'):
        try:
            self.com_object = self.app.com_object.GetBus(bus_type)
        except Exception as e:
            logger.error(f"❌ Error retrieving {bus_type} bus: {e}")
        finally:
            return self.com_object

    @property
    def active(self) -> bool:
        return self.com_object.Active

    @property
    def baudrate(self) -> int:
        return self.com_object.Baudrate()

    @baudrate.setter
    def baudrate(self, value: int):
        self.com_object.SetBaudrate(value)

    @property
    def name(self) -> str:
        return self.com_object.Name

    @name.setter
    def name(self, name: str):
        self.com_object.Name = name

    def get_signal(self, channel: int, message: str, signal: str) -> Signal:
        return Signal(self, channel, message, signal)

    def get_j1939_signal(self, channel: int, message: str, signal: str, source_address: int, destination_address: int) -> Signal:
        return Signal(self, channel, message, signal, source_address, destination_address)

    def get_bus_databases_info(self, bus: str = 'CAN') -> dict:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return {}
            databases_info = {}
            self.set_bus(bus_type)
            for db_obj in self.com_object.Databases:
                info = {
                    'full_name': getattr(db_obj, 'FullName', None),
                    'path': getattr(db_obj, 'Path', None),
                    'name': getattr(db_obj, 'Name', None),
                    'channel': getattr(db_obj, 'Channel', None),
                    'com_obj': db_obj,
                }
                databases_info[info['name']] = info
            logger.info(f'📜 {bus_type} bus databases information:')
            for db_name, db_info in databases_info.items():
                logger.info(f"    {db_name}:")
                for key, value in db_info.items():
                    logger.info(f"        {key}: {value}")
            return databases_info
        except Exception as e:
            logger.error(f"❌ Error retrieving {bus} bus databases information: {e}")
            return {}

    def get_bus_nodes_info(self, bus: str = 'CAN') -> dict:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return {}
            nodes_info = {}
            self.set_bus(bus_type)
            for node_obj in self.com_object.Nodes:
                info = {
                    'full_name': getattr(node_obj, 'FullName', None),
                    'path': getattr(node_obj, 'Path', None),
                    'name': getattr(node_obj, 'Name', None),
                    'active': getattr(node_obj, 'Active', None),
                    'com_obj': node_obj,
                }
                nodes_info[info['name']] = info
            logger.info(f'📜 {bus_type} bus nodes information:')
            for node_name, node_info in nodes_info.items():
                logger.info(f"    {node_name}:")
                for key, value in node_info.items():
                    logger.info(f"        {key}: {value}")
            return nodes_info
        except Exception as e:
            logger.error(f"❌ Error retrieving {bus} bus nodes information: {e}")
            return {}

    def get_signal_value(self, bus: str, channel: int, message: str, signal: str, raw_value: bool = False) -> Union[int, float, None]:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return None
            self.set_bus(bus_type)
            signal_obj = self.get_signal(channel, message, signal)
            value = signal_obj.raw_value if raw_value else signal_obj.value
            logger.info(f"🚦Signal({signal_obj.full_name}) value = {value}")
            return value
        except Exception as e:
            logger.error(f"❌ Error retrieving {bus} bus signal value: {e}")
            return None

    def set_signal_value(self, bus: str, channel: int, message: str, signal: str, value: Union[int, float], raw_value: bool = False) -> bool:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return False
            self.set_bus(bus_type)
            signal_obj = self.get_signal(channel, message, signal)
            if raw_value:
                signal_obj.raw_value = int(value)
            else:
                signal_obj.value = value
            logger.info(f"🚦Signal({signal_obj.full_name}) value set to {value}")
            return True
        except Exception as e:
            logger.error(f"❌ Error setting {bus} bus signal value: {e}")
            return False

    def get_signal_full_name(self, bus: str, channel: int, message: str, signal: str) -> Union[str, None]:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return None
            self.set_bus(bus_type)
            signal_obj = self.get_signal(channel, message, signal)
            full_name = signal_obj.full_name
            logger.info(f'🚦Signal full name = {full_name}')
            return full_name
        except Exception as e:
            logger.error(f"❌ Error retrieving {bus} bus signal full name: {e}")
            return None

    def check_signal_online(self, bus: str, channel: int, message: str, signal: str) -> bool:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return False
            self.set_bus(bus_type)
            signal_obj = self.get_signal(channel, message, signal)
            is_online = signal_obj.is_online
            logger.info(f'🚦Signal({signal_obj.full_name}) is online ?: {is_online} ({self.VALUE_TABLE_SIGNAL_IS_ONLINE[is_online]})')
            return is_online
        except Exception as e:
            logger.error(f"❌ Error checking {bus} bus signal online status: {e}")
            return False

    def check_signal_state(self, bus: str, channel: int, message: str, signal: str) -> int:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return -1
            self.set_bus(bus_type)
            signal_obj = self.get_signal(channel, message, signal)
            state = signal_obj.state
            logger.info(f'🚦Signal({signal_obj.full_name}) state: {state} ({self.VALUE_TABLE_SIGNAL_STATE[state]})')
            return state
        except Exception as e:
            logger.error(f"❌ Error checking {bus} bus signal state: {e}")
            return -1

    def get_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, raw_value=False) -> Union[float, int, None]:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return None
            self.set_bus(bus_type)
            signal_obj = self.get_j1939_signal(channel, message, signal, source_addr, dest_addr)
            signal_value = signal_obj.raw_value if raw_value else signal_obj.value
            logger.info(f'🚦J1939 Signal({signal_obj.full_name}) value = {signal_value}')
            return signal_value
        except Exception as e:
            logger.error(f"❌ Error retrieving J1939 bus signal value: {e}")
            return None

    def set_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, value: Union[float, int], raw_value: bool = False) -> bool:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return False
            self.set_bus(bus_type)
            signal_obj = self.get_j1939_signal(channel, message, signal, source_addr, dest_addr)
            if raw_value:
                signal_obj.raw_value = int(value)
            else:
                signal_obj.value = value
            logger.info(f'🚦J1939 Signal({signal_obj.full_name}) value set to {value}')
            return True
        except Exception as e:
            logger.error(f"❌ Error setting J1939 bus signal value: {e}")
            return False

    def get_j1939_signal_full_name(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> Union[str, None]:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return None
            self.set_bus(bus_type)
            signal_obj = self.get_j1939_signal(channel, message, signal, source_addr, dest_addr)
            full_name = signal_obj.full_name
            logger.info(f'🚦J1939 Signal full name = {full_name}')
            return full_name
        except Exception as e:
            logger.error(f"❌ Error retrieving J1939 bus signal full name: {e}")
            return None

    def check_j1939_signal_online(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> bool:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return False
            self.set_bus(bus_type)
            signal_obj = self.get_j1939_signal(channel, message, signal, source_addr, dest_addr)
            is_online = signal_obj.is_online
            logger.info(f'🚦J1939 Signal({signal_obj.full_name}) is online ?: {is_online} ({self.VALUE_TABLE_SIGNAL_IS_ONLINE[is_online]})')
            return is_online
        except Exception as e:
            logger.error(f"❌ Error checking J1939 bus signal online status: {e}")
            return False

    def check_j1939_signal_state(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> int:
        try:
            bus_type = bus.upper()
            if bus_type not in self.app.bus_types:
                logger.error(f"❌ Invalid bus type '{bus_type}'. Supported types: {', '.join(self.app.bus_types)}")
                return -1
            self.set_bus(bus_type)
            signal_obj = self.get_j1939_signal(channel, message, signal, source_addr, dest_addr)
            state = signal_obj.state
            logger.info(f'🚦J1939 Signal({signal_obj.full_name}) state: {state} ({self.VALUE_TABLE_SIGNAL_STATE[state]})')
            return state
        except Exception as e:
            logger.error(f"❌ Error checking J1939 bus signal state: {e}")
            return -1
