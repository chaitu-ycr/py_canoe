# import external modules here
import logging
import win32com.client

# import internal modules here


class Ui:
    """The UI object represents the user interface in CANoe.
    """
    def __init__(self, app_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.UI)
            self.write_window_com_obj = win32com.client.Dispatch(self.com_obj.Write)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe UI: {str(e)}')

    def get_command_availability(self, command: str) -> bool:
        """defines the availability of a command on the user interface.

        Args:
            command (str): The command. Currently only the values start and stop can be input for the command parameter.

        Returns:
            bool: The availability of the command: If the command is available True is returned. Otherwise, False is returned.
        """
        return self.com_obj.CommandEnabled(command)

    def set_command_availability(self, command: str, value: bool) -> None:
        """sets the availability of a command on the user interface.

        Args:
            command (str): The command. Currently only the values start and stop can be input for the command parameter.
            value (bool): A boolean value that indicates whether the command should be available. Possible values are: True: The command is available. False: The command is not available.
        """
        ce_obj = self.com_obj.CommandEnabled(command)
        ce_obj = value

    def activate_desktop(self, name: str) -> None:
        """Activates the desktop with the given name.

        Args:
            name (str): The name of the desktop to be activated.
        """
        self.com_obj.ActivateDesktop(name)

    def open_baudrate_dialog(self) -> None:
        """Configures the bus parameters.
        """
        self.com_obj.OpenBaudrateDialog()

    @property
    def get_write_window_text(self) -> str:
        """Gets the text contents of the Write window.

        Returns:
            str: The text content
        """
        return self.write_window_com_obj.Text

    def clear_write_window(self) -> None:
        """Clears the contents of the Write Window
        """
        self.write_window_com_obj.Clear()

    def copy_write_window_content(self) -> None:
        """Copies the contents of the Write Window to the clipboard.
        """
        self.write_window_com_obj.Copy()

    def disable_write_window_output_file(self, tab_index=None) -> None:
        """Disables logging of all outputs of the Write Window for the certain page.

        Args:
            tab_index (int, optional): The index of the tab, for which logging of the output is to be deactivated. Defaults to None.
        """
        if tab_index:
            self.write_window_com_obj.DisableOutputFile(tab_index)
        else:
            self.write_window_com_obj.DisableOutputFile()

    def enable_write_window_output_file(self, output_file: str, tab_index=None) -> None:
        """Enables logging of all outputs of the Write Window in the output file for the certain page.

        Args:
            output_file (str, optional): The complete path of the output file. Defaults to None.
            tab_index (int, optional): The index of the page, for which logging of the output is to be activated. Defaults to None.
        """
        if tab_index:
            self.write_window_com_obj.EnableOutputFile(output_file, tab_index)
        else:
            self.write_window_com_obj.EnableOutputFile(output_file)

    def output_text_in_write_window(self, text: str) -> None:
        """Outputs a line of text in the Write Window.

        Args:
            text (str): The text
        """
        self.write_window_com_obj.Output(text)
