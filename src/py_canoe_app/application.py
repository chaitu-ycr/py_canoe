# import external modules here
import sys
import logging
import pythoncom
import win32com.client

# import internal modules here
from .app_utils.bus import Bus
from .app_utils.capl import Capl
from .app_utils.configuration import Configuration
from .app_utils.environment import Environment
from .app_utils.measurement import Measurement
from .app_utils.networks import Networks
from .app_utils.performance import Performance
from .app_utils.simulation import Simulation
from .app_utils.system import System
from .app_utils.ui import Ui
from .app_utils.version import Version


class ApplicationEvents:
    """Handler for CANoe Application events"""

    @staticmethod
    def OnOpen(fullname):
        logging.getLogger('CANOE_LOG').debug(f'👉 canoe config ({fullname}) opened')
        Application.OPENED = True
        Application.CLOSED = False

    @staticmethod
    def OnQuit():
        logging.getLogger('CANOE_LOG').debug(f'👉 canoe closed')
        Application.OPENED = False
        Application.CLOSED = True


class Application:
    """Represents a CANoe application."""
    def __init__(self, user_capl_function_names: tuple, enable_app_events=False, enable_simulation=False):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            pythoncom.CoInitialize()
            self.user_capl_function_names = user_capl_function_names
            self.enable_simulation = enable_simulation
            self.com_obj = win32com.client.Dispatch('CANoe.Application')
            if enable_app_events:
                win32com.client.WithEvents(self.com_obj, ApplicationEvents)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe application: {str(e)}')
            sys.exit(1)

    @property
    def channel_mapping_name(self) -> str:
        return self.com_obj.ChannelMappingName

    @channel_mapping_name.setter
    def channel_mapping_name(self, name: str):
        self.com_obj.ChannelMappingName = name

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def path(self) -> str:
        return self.com_obj.Path

    @property
    def visible(self) -> bool:
        return self.com_obj.Visible

    @visible.setter
    def visible(self, visible: bool):
        self.com_obj.Visible = visible

    def new(self, auto_save=False, prompt_user=False) -> None:
        self.com_obj.New(auto_save, prompt_user)

    def open(self, path: str, auto_save=False, prompt_user=False) -> None:
        self.com_obj.Open(path, auto_save, prompt_user)

    def quit(self):
        self.com_obj.Quit()

    @property
    def bus(self) -> Bus:
        return Bus(self.com_obj)
    
    @property
    def capl(self) -> Capl:
        return Capl(self.com_obj)

    @property
    def configuration(self) -> Configuration:
        return Configuration(self.com_obj)
    
    @property
    def environment(self) -> Environment:
        return Environment(self.com_obj)

    @property
    def measurement(self) -> Measurement:
        return Measurement(self.com_obj, self.user_capl_function_names)
    
    @property
    def networks(self) -> Networks:
        return Networks(self.com_obj)
    
    @property
    def performance(self) -> Performance:
        return Performance(self.com_obj)

    @property
    def simulation(self) -> Simulation:
        return Simulation(self.com_obj, self.enable_simulation)

    @property
    def system(self) -> System:
        return System(self.com_obj)

    @property
    def ui(self) -> Ui:
        return Ui(self.com_obj)

    @property
    def version(self) -> Version:
        return Version(self.com_obj)
