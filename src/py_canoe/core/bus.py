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
    def __init__(self, app, bus_type: str = 'CAN'):
        self.bus_type = bus_type
        self.com_object = app.com_object.GetBus(bus_type)

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


def get_bus_databases_info(app, bus: str = 'CAN') -> dict:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return {}
        databases_info = {}
        for db_obj in app.com_object.GetBus(bus).Databases:
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
        logger.error(f"😡 Error retrieving {bus} bus databases information: {e}")
        return {}

def get_bus_nodes_info(app, bus: str = 'CAN') -> dict:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return {}
        nodes_info = {}
        for node_obj in app.com_object.GetBus(bus).Nodes:
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
        logger.error(f"😡 Error retrieving {bus} bus nodes information: {e}")
        return {}

def get_signal_value(app, bus: str, channel: int, message: str, signal: str, raw_value: bool = False) -> Union[int, float, None]:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return None
        signal_obj = app.com_object.GetBus(bus).GetSignal(channel, message, signal)
        value = signal_obj.RawValue if raw_value else signal_obj.Value
        logger.info(f"🚦Signal({bus}{channel}.{message}.{signal}) value = {value}")
        return value
    except Exception as e:
        logger.error(f"😡 Error retrieving {bus} bus signal value: {e}")
        return None

def set_signal_value(app, bus: str, channel: int, message: str, signal: str, value: Union[int, float], raw_value: bool = False) -> bool:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return False
        signal_obj = app.com_object.GetBus(bus).GetSignal(channel, message, signal)
        setattr(signal_obj, "RawValue" if raw_value else "Value", value)
        logger.info(f"🚦Signal({bus}{channel}.{message}.{signal}) value set to {value}")
        return True
    except Exception as e:
        logger.error(f"😡 Error setting {bus} bus signal value: {e}")
        return False

def get_signal_full_name(app, bus: str, channel: int, message: str, signal: str) -> Union[str, None]:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return None
        signal_obj = app.com_object.GetBus(bus).GetSignal(channel, message, signal)
        full_name = getattr(signal_obj, 'FullName', None)
        logger.info(f'🚦Signal full name = {full_name}')
        return full_name
    except Exception as e:
        logger.error(f"😡 Error retrieving {bus} bus signal full name: {e}")
        return None

def check_signal_online(app, bus: str, channel: int, message: str, signal: str) -> bool:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return False
        signal_obj = app.com_object.GetBus(bus).GetSignal(channel, message, signal)
        is_online = signal_obj.IsOnline
        logger.info(f'🚦Signal({bus}{channel}.{message}.{signal}) is online: {is_online}')
        return is_online
    except Exception as e:
        logger.error(f"😡 Error checking {bus} bus signal online status: {e}")
        return False

def check_signal_state(app, bus: str, channel: int, message: str, signal: str) -> int:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return -1
        signal_obj = app.com_object.GetBus(bus).GetSignal(channel, message, signal)
        state = signal_obj.State
        logger.info(f'🚦Signal({bus}{channel}.{message}.{signal}) state: {state}')
        return state
    except Exception as e:
        logger.error(f"😡 Error checking {bus} bus signal state: {e}")
        return -1

def get_j1939_signal_value(app, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, raw_value=False) -> Union[float, int, None]:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return None
        signal_obj = app.com_object.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
        signal_value = signal_obj.RawValue if raw_value else signal_obj.Value
        logger.info(f'🚦J1939 Signal({bus}{channel}.{message}.{signal}) value = {signal_value}')
        return signal_value
    except Exception as e:
        logger.error(f"😡 Error retrieving J1939 bus signal value: {e}")
        return None

def set_j1939_signal_value(app, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, value: Union[float, int], raw_value: bool = False) -> bool:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return False
        signal_obj = app.com_object.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
        setattr(signal_obj, "RawValue" if raw_value else "Value", value)
        logger.info(f'🚦J1939 Signal({bus}{channel}.{message}.{signal}) value set to {value}')
        return True
    except Exception as e:
        logger.error(f"😡 Error setting J1939 bus signal value: {e}")
        return False

def get_j1939_signal_full_name(app, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> Union[str, None]:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return None
        signal_obj = app.com_object.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
        full_name = signal_obj.FullName
        logger.info(f'🚦J1939 Signal full name = {full_name}')
        return full_name
    except Exception as e:
        logger.error(f"😡 Error retrieving J1939 bus signal full name: {e}")
        return None

def check_j1939_signal_online(app, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> bool:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return False
        signal_obj = app.com_object.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
        is_online = signal_obj.IsOnline
        logger.info(f'🚦J1939 Signal({bus}{channel}.{message}.{signal}) is online: {is_online}')
        return is_online
    except Exception as e:
        logger.error(f"😡 Error checking J1939 bus signal online status: {e}")
        return False

def check_j1939_signal_state(app, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> int:
    try:
        bus_type = bus.upper()
        if bus_type not in app.bus_type:
            logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(app.bus_type)}")
            return -1
        signal_obj = app.com_object.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
        state = signal_obj.State
        logger.info(f'🚦J1939 Signal({bus}{channel}.{message}.{signal}) state: {state}')
        return state
    except Exception as e:
        logger.error(f"😡 Error checking J1939 bus signal state: {e}")
        return -1
