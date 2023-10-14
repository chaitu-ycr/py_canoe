# Import Python Libraries here
import win32com.client


class Configuration:
    """The Configuration object represents the active configuration.
    """

    def __init__(self, app_obj) -> None:
        self.app_obj = app_obj
        self.log = self.app_obj.log
        self.conf_com_obj = win32com.client.Dispatch(self.app_obj.app_com_obj.Configuration)
        win32com.client.WithEvents(self.conf_com_obj, CanoeConfigurationEvents)

    @property
    def comment(self) -> str:
        """Gets the comment for the configuration.

        Returns:
            str: The comment.
        """
        return self.conf_com_obj.Comment

    @comment.setter
    def comment(self, text: str) -> None:
        """Defines the comment for the configuration.

        Args:
            text (str): The comment.
        """
        self.conf_com_obj.Comment = text
        self.log.info(f'configuration comment set to {text}.')

    @property
    def fdx_enabled(self) -> int:
        """Enables/Disables value of FDX protocol.

        Returns:
            int: The activation state of the FDX protocol. 0: FDX protocol is deactivated. 1: FDX protocol is activated.
        """
        return self.conf_com_obj.FDXEnabled

    @fdx_enabled.setter
    def fdx_enabled(self, enabled: int) -> None:
        """Enables/Disables the FDX protocol.

        Args:
            enabled (int): The activation state of the FDX protocol. 0: deactivate FDX protocol. ≠0: activate FDX protocol.
        """
        self.conf_com_obj.FDXEnabled = enabled
        self.log.info(f'FDX protocol set to {enabled}.')

    @property
    def full_name(self) -> str:
        """gets the complete path of the configuration.

        Returns:
            str: complete path of the configuration.
        """
        return self.conf_com_obj.FullName

    @full_name.setter
    def full_name(self, full_name: str) -> None:
        """sets the complete path of the configuration.

        Args:
            full_name (str): The new complete path of the configuration.
        """
        self.conf_com_obj.FullName = full_name
        self.log.info(f'complete path of the configuration set to {full_name}.')

    @property
    def mode(self) -> int:
        """returns whether the Online mode or the Offline mode is active.

        Returns:
            int: The currently active mode.
        """
        return self.conf_com_obj.Mode

    @mode.setter
    def mode(self, mode: int) -> None:
        """sets the Online mode or the Offline mode to active.

        Args:
            mode (int): The active mode; valid values are: 0-Online mode is activated. 1-Offline mode is activated.
        """
        self.conf_com_obj.Mode = mode
        self.log.info(f'offline/online mode set to {mode}.')

    @property
    def modified(self) -> bool:
        """Returns information on whether the current configuration was modified since the time it was loaded or created, or sets this property.
        This property determines whether the user is prompted to save when another configuration is loaded.

        Returns:
            bool: The current value of the property.
        """
        return self.conf_com_obj.Modified

    @modified.setter
    def modified(self, modified: bool) -> None:
        """sets Modified property to flase/true.

        Args:
            modified (bool): Value to be assigned to the Modified property.
        """
        self.conf_com_obj.Modified = modified
        self.log.info(f'configuration modified property set to {modified}.')

    @property
    def name(self) -> str:
        """Returns the name of the configuration.

        Returns:
            str: The name of the currently loaded configuration.
        """
        return self.conf_com_obj.Name

    @property
    def path(self) -> str:
        """returns the path of the configuration, depending on the actual configuration.

        Returns:
            str: The path of the currently loaded configuration.
        """
        return self.conf_com_obj.Path

    @property
    def read_only(self) -> bool:
        """Indicates whether the configuration is write protected.

        Returns:
            bool: If the object is write protected True is returned; otherwise False is returned.
        """
        return self.conf_com_obj.ReadOnly

    @property
    def saved(self) -> bool:
        """Indicates whether changes to the configuration have already been saved.

        Returns:
            bool: If changes were made to the configuration and they have not been saved yet, False is returned; otherwise True is returned.
        """
        return self.conf_com_obj.Saved

    def compile_and_verify(self):
        """Compiles all CAPL test modules and verifies all XML test modules.
        All test modules in the Simulation Setup and in the Test Setup are taken into consideration.
        """
        self.conf_com_obj.CompileAndVerify()
        self.log.info(f'Compiled all test modules in the Simulation Setup and in the Test Setup.')

    def save(self, path='', prompt_user=False):
        """Saves the configuration.

        Args:
            path (str): The complete file name. If no path is specified, the configuration is saved under its current name. If it is not saved yet, the user will be prompted for a name.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations.
        """
        if not self.saved:
            if path == '':
                self.conf_com_obj.Save()
            else:
                self.conf_com_obj.Save(path, prompt_user)
            self.log.info(f'Saved configuration({path}).')
        else:
            self.log.info('CANoe Configuration already in saved state.')
        return self.saved

    def save_as(self, path: str, major: str, minor: str, prompt_user: bool):
        """Saves the configuration as a different CANoe version

        Args:
            path (str): The complete path.
            major (str): The major version number of the target version, e.g. 10 for CANoe 10.1.
            minor (str): The minor version number of the target version, e.g. 1 for CANoe 10.1
            prompt_user (bool): A boolean value that defines whether the user should interact in error situations.
        """
        self.conf_com_obj.SaveAs(path, major, minor, prompt_user)
        self.log.info(f'Saved configuration as {path}.')


class CanoeConfigurationEvents:
    """Handler for CANoe Configuration events"""

    @staticmethod
    def OnClose():
        """Occurs when the configuration is closed.
        """
        print('configuration OnClose event triggered.')

    @staticmethod
    def OnSystemVariablesDefinitionChanged():
        """Occurs when system variable definitions are added, changed or removed.
        """
        print('configuration OnSystemVariablesDefinitionChanged event triggered.')
