# import external modules here
import os
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
    """Represents a CANoe application."""
    def __init__(self, enable_app_events=False, enable_simulation=False):
        self.__log = logging.getLogger('CANOE_LOG')
        pythoncom.CoInitialize()
        try:
            self.com_obj = win32com.client.Dispatch('CANoe.Application')
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe application: {str(e)}')
            sys.exit(1)

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

    def open(self, path: str, auto_save=False, prompt_user=False) -> None:
        """Loads a configuration.

        Args:
            path (str): The complete path for the configuration.
            auto_save (bool, optional): A boolean value that indicates whether the active configuration should be saved if it has been changed. Defaults to False.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations. Defaults to False.

        Raises:
            FileNotFoundError: error when canoe config file not available in pc.
        """
        self.com_obj.Open(path, auto_save, prompt_user)

    def quit(self):
        """Quits the application."""
        self.com_obj.Quit()

    @property
    def bus(self) -> Bus:
        """Returns the Bus object.

        Returns:
            Bus: The Bus object.
        """
        return Bus(self.com_obj)

    @property
    def configuration(self) -> Configuration:
        """Returns the Configuration object.

        Returns:
            Configuration: The Configuration object.
        """
        return Configuration(self.com_obj)

    @property
    def measurement(self) -> Measurement:
        """Returns the Measurement object.

        Returns:
            Measurement: The Measurement object.
        """
        return Measurement(self.com_obj)

    @property
    def system(self) -> System:
        """Returns the System object.

        Returns:
            System: The System object.
        """
        return System(self.com_obj)

    @property
    def ui(self) -> Ui:
        """Returns the Ui object.

        Returns:
            Ui: The Ui object.
        """
        return Ui(self.com_obj)

    @property
    def version(self) -> Version:
        """Returns the Version object.

        Returns:
            Version: The Version object.
        """
        return Version(self.com_obj)
