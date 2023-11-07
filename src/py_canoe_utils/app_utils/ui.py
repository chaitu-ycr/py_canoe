# Import Python Libraries here
import logging
import win32com.client

logger_inst = logging.getLogger('CANOE_LOG')


class Ui:
    """The UI object represents the user interface in CANoe.
    """

    def __init__(self, app_com_obj):
        self.__log = logger_inst
        self.com_obj = win32com.client.Dispatch(app_com_obj.UI)
        self.write = Write(self)

    def get_command_availability(self, command: str) -> bool:
        """defines the availability of a command on the user interface.

        Args:
            command (str): The command. If no command is entered the function acts on all commands that can be influenced via this interface.

        Returns:
            bool: The availability of the command: If the command is available True is returned. Otherwise, False is returned.
        """
        return self.com_obj.CommandEnabled(command)

    def set_command_availability(self, command: str, value: bool) -> None:
        """sets the availability of a command on the user interface.

        Args:
            command (str): The command. If no command is entered the function acts on all commands that can be influenced via this interface.
            value (bool): A boolean value that indicates whether the command should be available. Possible values are: True: The command is available. False: The command is not available.
        """
        ce_obj = self.com_obj.CommandEnabled(command)
        ce_obj = value
        self.__log.info(f'enabled command {command}.')

    def activate_desktop(self, name: str) -> None:
        """Activates the desktop with the given name.

        Args:
            name (str): The name of the desktop to be activated.
        """
        self.com_obj.ActivateDesktop(name)
        self.__log.info(f'Activated the desktop with the given name({name}.')

    def open_baudrate_dialog(self) -> None:
        """Configures the bus parameters.
        """
        self.com_obj.OpenBaudrateDialog()
        self.__log.info(f'baudrate dialog opened. Configure the bus parameters.')


class Write:
    """The Write object represents the Write Window in CANoe.
    It is part of the user interface.
    """

    def __init__(self, ui_obj):
        self.__log = logger_inst
        self.com_obj = win32com.client.Dispatch(ui_obj.com_obj.Write)

    @property
    def text(self) -> str:
        """Gets the text contents of the Write window.

        Returns:
            str: The text content
        """
        return self.com_obj.Text

    def clear(self) -> None:
        """Clears the contents of the Write Window
        """
        self.com_obj.Clear()
        self.__log.info(f'Cleared the contents of the Write Window.')

    def copy(self) -> None:
        """Copies the contents of the Write Window to the clipboard.
        """
        self.com_obj.Copy()
        self.__log.info(f'Copied the contents of the Write Window to the clipboard.')

    def disable_output_file(self, tab_index=None) -> None:
        """Disables logging of all outputs of the Write Window for the certain page.

        Args:
            tab_index (int, optional): The index of the tab, for which logging of the output is to be deactivated. Defaults to None.
        """
        if tab_index:
            self.com_obj.DisableOutputFile(tab_index)
        else:
            self.com_obj.DisableOutputFile()
        self.__log.info(f'Disabled logging of outputs of the Write Window. tab_index={tab_index}')

    def enable_output_file(self, output_file: str, tab_index=None) -> None:
        """Enables logging of all outputs of the Write Window in the output file for the certain page.

        Args:
            output_file (str, optional): The complete path of the output file. Defaults to None.
            tab_index (int, optional): The index of the page, for which logging of the output is to be activated. Defaults to None.
        """
        if tab_index:
            self.com_obj.EnableOutputFile(output_file, tab_index)
        else:
            self.com_obj.EnableOutputFile(output_file)
        self.__log.info(
            f'Enabled logging of outputs of the Write Window. output_file={output_file} and tab_index={tab_index}')

    def output(self, text: str) -> None:
        """Outputs a line of text in the Write Window.

        Args:
            text (str): The text
        """
        self.com_obj.Output(text)
        self.__log.info(f'Outputed {text} in the Write Window.')
