# Import Python Libraries here
import logging
import os
import win32com.client

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
        """Occurs when a configuration is opened.

        Args:
            fullname (str): The complete file name of the loaded configuration.
        """
        print(f'canoe config ({fullname}) opened')
        Application.OPENED = True
        Application.CLOSED = False

    @staticmethod
    def OnQuit():
        """Occurs when CANoe is quit
        """
        print('canoe closed')
        Application.OPENED = False
        Application.CLOSED = True


class Application:
    """The Application object represents the CANoe application.
    """
    OPENED = False
    CLOSED = False

    def __init__(self, user_capl_function_names: tuple, enable_app_events=False, enable_simulation=False):
        self.log = logging.getLogger('CANOE_LOG')
        self.user_capl_function_names = user_capl_function_names
        self.enable_app_events = enable_app_events
        self.enable_simulation = enable_simulation
        self.com_obj = win32com.client.Dispatch('CANoe.Application')
        self.bus: Bus
        self.capl: Capl
        self.configuration: Configuration
        self.environment: Environment
        self.measurement: Measurement
        self.networks: Networks
        self.performance: Performance
        self.simulation: Simulation
        self.system: System
        self.ui: Ui
        self.version: Version
        self.__print_application_info()

    def __print_application_info(self):
        cav = self.com_obj.Version
        self.log.info(f'Dispatched Vector CANoe Application {cav.major}.{cav.minor}.{cav.Build}')

    def __initialise_application_child_objects(self):
        self.bus = Bus(self.com_obj)
        self.capl = Capl(self.com_obj)
        self.configuration = Configuration(self.com_obj)
        self.environment = Environment(self.com_obj)
        self.networks = Networks(self.com_obj)
        self.performance = Performance(self.com_obj)
        self.system = System(self.com_obj)
        self.ui = Ui(self.com_obj)
        self.version = Version(self.com_obj)
        if self.enable_app_events:
            win32com.client.WithEvents(self.com_obj, ApplicationEvents)
        if self.enable_simulation:
            self.simulation = Simulation(self.com_obj)
        self.measurement = Measurement(self.com_obj, self.user_capl_function_names)

    @property
    def channel_mapping_name(self) -> str:
        """get the application name which is used to map application channels to real existing Vector hardware interface channels.

        Returns:
            str: The application name
        """
        return self.com_obj.ChannelMappingName

    @channel_mapping_name.setter
    def channel_mapping_name(self, name: str):
        """set the application name which is used to map application channels to real existing Vector hardware interface channels.

        Args:
            name (str): The application name used to map the channels.
        """
        self.com_obj.ChannelMappingName = name

    @property
    def full_name(self) -> str:
        """determines the complete path of the CANoe application.

        Returns:
            str: location where CANoe is installed.
        """
        return self.com_obj.FullName

    @property
    def name(self) -> str:
        """Returns the name of the CANoe application.

        Returns:
            str: name of the CANoe application.
        """
        return self.com_obj.Name

    @property
    def path(self) -> str:
        """Returns the Path of the CANoe application.

        Returns:
            str: Path of the CANoe application.
        """
        return self.com_obj.Path

    @property
    def visible(self) -> bool:
        """Returns whether the CANoe main window is visible or is only displayed by a tray icon.

        Returns:
            bool: A boolean value indicating whether the CANoe main window is visible..
        """
        return self.com_obj.Visible

    @visible.setter
    def visible(self, visible: bool):
        """Defines whether the CANoe main window is visible or is only displayed by a tray icon.

        Args:
            visible (bool): A boolean value indicating whether the CANoe main window is to be visible.
        """
        self.com_obj.Visible = visible

    def new(self, auto_save=False, prompt_user=False) -> None:
        """Creates a new configuration.

        Args:
            auto_save (bool, optional): A boolean value that indicates whether the active configuration should be saved if it has been changed. Defaults to False.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations. Defaults to False.
        """
        self.com_obj.New(auto_save, prompt_user)
        self.log.info('created a new configuration...')

    def open(self, path: str, auto_save=False, prompt_user=False) -> None:
        """Loads a configuration.

        Args:
            path (str): The complete path for the configuration.
            auto_save (bool, optional): A boolean value that indicates whether the active configuration should be saved if it has been changed. Defaults to False.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations. Defaults to False.

        Raises:
            FileNotFoundError: _description_
        """
        if not auto_save:
            self.com_obj.Configuration.Modified = False
            self.log.info(f'CANoe cfg "Modified" parameter set to False to avoid error.')
        if os.path.isfile(path):
            self.log.info(f'CANoe cfg "{path}" found.')
            self.com_obj.Open(path, auto_save, prompt_user)
            self.__initialise_application_child_objects()
            self.log.info(f'loaded CANoe config "{path}"')

        else:
            self.log.info(f'CANoe cfg "{path}" not found.')
            raise FileNotFoundError(f'CANoe cfg file "{path}" not found!')

    def quit(self):
        """Quits the application.
        """
        self.com_obj.Quit()
        self.log.info('CANoe Application Closed.')
