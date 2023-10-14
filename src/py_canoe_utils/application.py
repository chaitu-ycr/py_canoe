# Import Python Libraries here
import os


class Application:
    """The Application object represents the CANoe application.
    """
    OPENED = False
    CLOSED = False

    def __init__(self, log_obj) -> None:
        self.app_com_obj = None
        self.log = log_obj

    @property
    def channel_mapping_name(self) -> str:
        """get the application name which is used to map application channels to real existing Vector hardware interface channels.

        Returns:
            str: The application name
        """
        return self.app_com_obj.ChannelMappingName

    @channel_mapping_name.setter
    def channel_mapping_name(self, name: str):
        """set the application name which is used to map application channels to real existing Vector hardware interface channels.

        Args:
            name (str): The application name used to map the channels.
        """
        self.app_com_obj.ChannelMappingName = name

    @property
    def full_name(self) -> str:
        """determines the complete path of the CANoe application.

        Returns:
            str: location where CANoe is installed.
        """
        return self.app_com_obj.FullName

    @property
    def name(self) -> str:
        """Returns the name of the CANoe application.

        Returns:
            str: name of the CANoe application.
        """
        return self.app_com_obj.Name

    @property
    def path(self) -> str:
        """Returns the Path of the CANoe application.

        Returns:
            str: Path of the CANoe application.
        """
        return self.app_com_obj.Path

    @property
    def visible(self) -> bool:
        """Returns whether the CANoe main window is visible or is only displayed by a tray icon.

        Returns:
            bool: A boolean value indicating whether the CANoe main window is visible..
        """
        return self.app_com_obj.Visible

    @visible.setter
    def visible(self, visible: bool):
        """Defines whether the CANoe main window is visible or is only displayed by a tray icon.

        Args:
            visible (bool): A boolean value indicating whether the CANoe main window is to be visible.
        """
        self.app_com_obj.Visible = visible

    def new(self, auto_save=False, prompt_user=False) -> None:
        """Creates a new configuration.

        Args:
            auto_save (bool, optional): A boolean value that indicates whether the active configuration should be saved if it has been changed. Defaults to False.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations. Defaults to False.
        """
        self.app_com_obj.New(auto_save, prompt_user)
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
        if os.path.isfile(path):
            self.log.info(f'CANoe cfg "{path}" found.')
            self.app_com_obj.Open(path, auto_save, prompt_user)
            self.log.info(f'loaded CANoe config "{path}"')
        else:
            self.log.info(f'CANoe cfg "{path}" not found.')
            raise FileNotFoundError(f'CANoe cfg file "{path}" not found!')

    def quit(self):
        """Quits the application.
        """
        self.app_com_obj.Quit()
        self.log.info('CANoe Application Closed.')


class ApplicationEvents:
    """Handler for CANoe Application events"""

    @staticmethod
    def OnOpen():
        """Occurs when a configuration is opened.
        """
        print('canoe opened')
        Application.OPENED = True
        Application.CLOSED = False

    @staticmethod
    def OnQuit():
        """Occurs when CANoe is quit
        """
        print('canoe closed')
        Application.OPENED = False
        Application.CLOSED = True
