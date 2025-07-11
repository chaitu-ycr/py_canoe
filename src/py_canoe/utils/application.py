import logging
import win32com.client

from py_canoe.utils.bus import Bus
from py_canoe.utils.capl import Capl
from py_canoe.utils.configuration import Configuration
from py_canoe.utils.environment import Environment
from py_canoe.utils.measurement import Measurement
from py_canoe.utils.networks import Networks
from py_canoe.utils.performance import Performance
from py_canoe.utils.simulation import Simulation
from py_canoe.utils.system import System
from py_canoe.utils.ui import Ui
from py_canoe.utils.version import Version

logging.getLogger('py_canoe')

class ApplicationEvents:
    @staticmethod
    def OnOpen(fullname):
        logging.info(f"Opened CANoe Configuration: {fullname}")

    @staticmethod
    def OnQuit():
        logging.info("CANoe Application Quit")

class Application:
    def __init__(self, enable_events: bool = True):
        self.com_object = win32com.client.Dispatch("CANoe.Application")
        if enable_events:
            win32com.client.WithEvents(self.com_object, ApplicationEvents)

    def bus(self, bus_type: str = 'CAN') -> Bus:
        if bus_type == 'CAN':
            return Bus(self.com_object, 'CAN')
        elif bus_type == 'LIN':
            return Bus(self.com_object, 'LIN')
        elif bus_type == 'FlexRay':
            return Bus(self.com_object, 'FlexRay')
        elif bus_type == 'AFDX':
            return Bus(self.com_object, 'AFDX')
        elif bus_type == 'Ethernet':
            return Bus(self.com_object, 'Ethernet')
        else:
            raise ValueError(f"Unsupported bus type: {bus_type}")

    @property
    def capl(self) -> Capl:
        return Capl(self)

    @property
    def channel_mapping_name(self) -> str:
        """The application name which is used to map application channels to real existing Vector hardware interface channels."""
        return self.com_object.ChannelMappingName

    @channel_mapping_name.setter
    def channel_mapping_name(self, name: str):
        """Set the application name which is used to map application channels to real existing Vector hardware interface channels."""
        self.com_object.ChannelMappingName = name

    @property
    def configuration(self) -> Configuration:
        return Configuration(self)

    @property
    def environment(self) -> Environment:
        return Environment(self)

    @property
    def full_name(self) -> str:
        return self.com_object.FullName

    @property
    def measurement(self) -> Measurement:
        return Measurement(self)

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def networks(self) -> Networks:
        return Networks(self)

    @property
    def path(self) -> str:
        return self.com_object.Path

    @property
    def performance(self) -> Performance:
        return Performance(self)

    @property
    def simulation(self) -> Simulation:
        return Simulation(self)

    @property
    def system(self) -> System:
        return System(self)

    @property
    def ui(self) -> Ui:
        return Ui(self)

    @property
    def version(self) -> Version:
        return Version(self)

    @property
    def visible(self) -> bool:
        return self.com_object.Visible

    @visible.setter
    def visible(self, visible: bool = True):
        self.com_object.Visible = visible

    def new(self, auto_save: bool = True, prompt_user: bool = False):
        self.com_object.New(auto_save, prompt_user)

    def open(self, path: str, auto_save: bool = True, prompt_user: bool = False):
        self.com_object.Open(path, auto_save, prompt_user)

    def quit(self):
        self.com_object.Quit()
