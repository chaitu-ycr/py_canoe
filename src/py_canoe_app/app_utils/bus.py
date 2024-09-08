# import external modules here
import logging
import win32com.client

# import internal modules here


class Bus:
    """The Bus object represents a bus of the CANoe application."""
    def __init__(self, app_com_obj):
        self.__log = logging.getLogger('CANOE_LOG')
        self.app_com_obj = app_com_obj
        self.com_obj = win32com.client.Dispatch(app_com_obj.Bus)

    def get_signal(self, bus: str, channel: int, message: str, signal: str) -> object:
        try:
            return SignalProperties(self.app_com_obj.GetBus(bus).GetSignal(channel, message, signal))
        except Exception as e:
            self.__log.error(f'Error getting signal: {str(e)}')

    def get_j1939_signal(self, bus: str, channel: int, message: str, signal: str, source_address: int, destination_address: int) -> object:
        try:
            return SignalProperties(self.app_com_obj.GetBus(bus).GetJ1939Signal(channel, message, signal, source_address, destination_address))
        except Exception as e:
            self.__log.error(f'Error getting J1939 signal: {str(e)}')

    def active(self, bus: str) -> bool:
        try:
            return self.app_com_obj.GetBus(bus).Active
        except Exception as e:
            self.__log.error(f'Error getting bus active: {str(e)}')

    def baudrate(self, bus: str, channel: int) -> float:
        try:
            return self.app_com_obj.GetBus(bus).Baudrate(channel)
        except Exception as e:
            self.__log.error(f'Error getting baudrate: {str(e)}')

    def name(self, bus: str) -> str:
        try:
            return self.app_com_obj.GetBus(bus).Name
        except Exception as e:
            self.__log.error(f'Error getting bus name: {str(e)}')

    @property
    def channels(self) -> object:
        return Channels(self.com_obj)

    @property
    def databases(self) -> object:
        return Databases(self.com_obj)

    @property
    def nodes(self) -> object:
        return Nodes(self.com_obj)

    @property
    def ports(self) -> object:
        return Ports(self.com_obj)

    @property
    def ports_of_channel(self) -> object:
        return Ports(self.com_obj)

    @property
    def replay_collection(self) -> object:
        return ReplayCollection(self.com_obj)

    @property
    def security_configuration(self) -> object:
        return SecurityConfiguration(self.com_obj)


class SignalProperties:
    def __init__(self, signal_com_obj):
        self.com_obj = signal_com_obj

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @property
    def is_online(self) -> bool:
        return self.com_obj.IsOnline

    @property
    def raw_value(self) -> int:
        return self.com_obj.RawValue

    @raw_value.setter
    def raw_value(self, value: int):
        self.com_obj.RawValue = value

    @property
    def state(self) -> int:
        return self.com_obj.State

    @property
    def value(self):
        return self.com_obj.Value

    @value.setter
    def value(self, value):
        self.com_obj.Value = value


class Channels:
    """The Channels object represents the channel configuration in the Simulation Setup of the CANoe application."""
    def __init__(self, bus_com_obj):
        self.com_obj = win32com.client.Dispatch(bus_com_obj.Channels)

    @property
    def count(self) -> int:
        return self.com_obj.Count


    def channel(self, index: int) -> object:
        return Channel(self.com_obj, index)


class Channel:
    """The Channel object represents a channel in the Simulation Setup of the CANoe application."""
    def __init__(self, channels_com_obj, index: int):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(channels_com_obj.Item(index))
        except Exception as e:
            self.__log.error(f'Error initializing channel: {str(e)}')

    @property
    def bus_type(self) -> str:
        return self.com_obj.BusType

    @property
    def controller(self) -> str:
        return self.com_obj.Controller

    @property
    def number(self) -> int:
        return self.com_obj.Number


class Databases:
    """The Databases object represents the assigned databases of CANoe."""
    def __init__(self, bus_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(bus_com_obj.Databases)
        except Exception as e:
            self.__log.error(f'Error initializing databases: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def database(self, index: int) -> object:
        return Database(self.com_obj, index)

    def add(self, name: str) -> None:
        try:
            self.com_obj.Add(name)
        except Exception as e:
            self.__log.error(f'Error adding database: {str(e)}')

    def add_network(self, database_name: str, network_name: str) -> None:
        try:
            self.com_obj.AddNetwork(database_name, network_name)
        except Exception as e:
            self.__log.error(f'Error adding database of network: {str(e)}')

    def remove(self, index: int) -> None:
        try:
            self.com_obj.Remove(index)
        except Exception as e:
            self.__log.error(f'Error removing database: {str(e)}')


class Database:
    def __init__(self, databases_com_obj, index: int):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(databases_com_obj.Item(index))
        except Exception as e:
            self.__log.error(f'Error initializing database: {str(e)}')

    @property
    def channel(self) -> int:
        return self.com_obj.Channel

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def path(self) -> str:
        return self.com_obj.Path


class Nodes:
    """"Node" describes a Program Node, e.g. CAPL/C# nodes of the Simulation Setup, or Application Models of the System and Communication Setup."""
    def __init__(self, bus_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(bus_com_obj.Nodes)
        except Exception as e:
            self.__log.error(f'Error initializing nodes: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def node(self, index: int) -> object:
        return Node(self.com_obj, index)

    def add(self, name: str) -> object:
        try:
            return self.com_obj.Add(name)
        except Exception as e:
            self.__log.error(f'Error adding node: {str(e)}')

    def add_test_module_ex(self, name: str, type: int) -> object:
        try:
            return self.com_obj.AddTestModuleEx(name, type)
        except Exception as e:
            self.__log.error(f'Error adding test module: {str(e)}')

    def add_with_tile(self, name: str) -> object:
        try:
            return self.com_obj.AddWithTile(name)
        except Exception as e:
            self.__log.error(f'Error adding node with tile: {str(e)}')

    def remove(self, index: int) -> None:
        try:
            self.com_obj.Remove(index)
        except Exception as e:
            self.__log.error(f'Error removing node: {str(e)}')


class Node:
    def __init__(self, nodes_com_obj, index: int):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(nodes_com_obj.Item(index))
        except Exception as e:
            self.__log.error(f'Error initializing node: {str(e)}')

    @property
    def active(self) -> bool:
        return self.com_obj.Active

    @property
    def attached_buses(self) -> object:
        return self.com_obj.AttachedBuses

    @property
    def drift_jitter_max(self) -> float:
        return self.com_obj.DriftJitterMax

    @property
    def drift_jitter_min(self) -> float:
        return self.com_obj.DriftJitterMin

    @property
    def drift_jitter_mode(self) -> int:
        return self.com_obj.DriftJitterMode

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @property
    def inputs(self) -> object:
        return self.com_obj.Inputs

    @property
    def is_gateway(self) -> bool:
        return self.com_obj.IsGateway

    @property
    def modules(self) -> object:
        return self.com_obj.Modules

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def outputs(self) -> object:
        return self.com_obj.Outputs

    @property
    def path(self) -> str:
        return self.com_obj.Path

    @property
    def tcp_ip_stack_setting(self) -> object:
        return self.com_obj.TcpIpStackSetting

    @property
    def test_module(self) -> bool:
        return self.com_obj.TestModule

    def attach_bus(self, bus: object) -> None:
        try:
            self.com_obj.AttachBus(bus)
        except Exception as e:
            self.__log.error(f'Error attaching bus: {str(e)}')

    def detach_bus(self, bus: object) -> None:
        try:
            self.com_obj.DetachBus(bus)
        except Exception as e:
            self.__log.error(f'Error detaching bus: {str(e)}')

    def is_bus_attached(self, bus: object) -> bool:
        try:
            return self.com_obj.IsBusAttached(bus)
        except Exception as e:
            self.__log.error(f'Error checking if bus is attached: {str(e)}')


class Ports:
    """The Ports object represents all ports of a specific Ethernet bus in Network-based Access mode as a collection."""
    def __init__(self, bus_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(bus_com_obj.Ports)
        except Exception as e:
            self.__log.error(f'Error initializing ports: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count

    @property
    def is_port_based_config(self) -> bool:
        return self.com_obj.IsPortBasedConfig

    @property
    def is_switched_network(self) -> bool:
        return self.com_obj.IsSwitchedNetwork

    @property
    def network_name(self) -> str:
        return self.com_obj.NetworkName

    @property
    def ports_are_simulated(self) -> bool:
        return self.com_obj.PortsAreSimulated

    def port(self, index: int) -> object:
        return Port(self.com_obj, index)

    def add(self, port_name: str, segment_name: str) -> object:
        try:
            return self.com_obj.Add(port_name, segment_name)
        except Exception as e:
            self.__log.error(f'Error adding port: {str(e)}')

    def add_mp(self, port_name: str) -> object:
        try:
            return self.com_obj.AddMP(port_name)
        except Exception as e:
            self.__log.error(f'Error adding port: {str(e)}')

    def remove(self, index: int) -> None:
        try:
            self.com_obj.Remove(index)
        except Exception as e:
            self.__log.error(f'Error removing port: {str(e)}')


class Port:
    """The Port object represents a specific port of a CANoe configuration.
    Ports are the access points for applications on the network. They can be used for simulation (read/write access) or for measurement (read access).
    """
    def __init__(self, ports_com_obj, index: int):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(ports_com_obj.Item(index))
        except Exception as e:
            self.__log.error(f'Error initializing port: {str(e)}')

    @property
    def is_active(self) -> bool:
        return self.com_obj.IsActive

    @is_active.setter
    def is_active(self, value: bool):
        self.com_obj.IsActive = value

    @property
    def is_simulated(self) -> bool:
        return self.com_obj.IsSimulated

    @is_simulated.setter
    def is_simulated(self, value: bool):
        self.com_obj.IsSimulated = value

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def segment_name(self) -> str:
        return self.com_obj.SegmentName

    @segment_name.setter
    def segment_name(self, value: str):
        self.com_obj.SegmentName = value


class ReplayCollection:
    """The ReplayCollection object represents the Replay Blocks of the CANoe application."""
    def __init__(self, bus_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(bus_com_obj.ReplayCollection)
        except Exception as e:
            self.__log.error(f'Error initializing replay collection: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def replay_block(self, index: int) -> object:
        return ReplayBlock(self.com_obj, index)

    def add(self, bus_type: int, name: str) -> object:
        try:
            return self.com_obj.Add(bus_type, name)
        except Exception as e:
            self.__log.error(f'Error adding replay: {str(e)}')

    def remove(self, index: int) -> None:
        try:
            self.com_obj.Remove(index)
        except Exception as e:
            self.__log.error(f'Error removing replay: {str(e)}')


class ReplayBlock:
    """The ReplayBlock object represents a Replay Block of the CANoe application."""
    def __init__(self, replay_collection_com_obj, index: int):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(replay_collection_com_obj.Item(index))
        except Exception as e:
            self.__log.error(f'Error initializing replay block: {str(e)}')

    @property
    def enabled(self) -> bool:
        return self.com_obj.Enabled

    @enabled.setter
    def enabled(self, value: bool):
        self.com_obj.Enabled = value

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @name.setter
    def name(self, name: str):
        self.com_obj.Name = name

    @property
    def path(self) -> str:
        return self.com_obj.Path

    @path.setter
    def path(self, value: str):
        self.com_obj.Path = value

    def start(self) -> None:
        try:
            self.com_obj.Start()
        except Exception as e:
            self.__log.error(f'Error starting replay block: {str(e)}')

    def stop(self) -> None:
        try:
            self.com_obj.Stop()
        except Exception as e:
            self.__log.error(f'Error stopping replay block: {str(e)}')


class SecurityConfiguration:
    """The SecurityConfiguration object represents a security profile assignment to a network, TCP stack or observer."""
    def __init__(self, bus_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(bus_com_obj.SecurityConfiguration)
        except Exception as e:
            self.__log.error(f'Error initializing security configuration: {str(e)}')

    def security_active(self, value: bool):
        self.com_obj.SecurityActive = value

    @property
    def security_profile(self) -> int:
        return self.com_obj.SecurityProfile

    @security_profile.setter
    def security_profile(self, value: int):
        self.com_obj.SecurityProfile = value