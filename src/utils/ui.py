# Import Python Libraries here
import win32com.client

class Ui:
    """The UI object represents the user interface in CANoe.
    """
    def __init__(self, app) -> None:
        self.app = app
        self.log = self.app.log
        self.ui_com_obj = win32com.client.Dispatch(self.app.app_com_obj.UI)
    
    @property
    def command_enabled(self, command: str) -> bool:
        """defines the availability of a command on the user interface.

        Args:
            command (str): The command. If no command is entered the function acts on all commands that can be influenced via this interface.

        Returns:
            bool: The availability of the command: If the command is available True is returned. Otherwise False is returned.
        """
        return self.ui_com_obj.CommandEnabled(command)
    
    @command_enabled.setter
    def command_enabled(self, command: str, value: bool) -> None:
        """sets the availability of a command on the user interface.

        Args:
            command (str): The command. If no command is entered the function acts on all commands that can be influenced via this interface.
            value (bool): A boolean value that indicates whether the command should be available. Possible values are: True: The command is available. False: The command is not available.
        """
        ce_obj = self.ui_com_obj.CommandEnabled(command)
        ce_obj = value
        self.log.info(f'enabled command {command}.')

    def activate_desktop(self, name: str) -> None:
        """Activates the desktop with the given name.

        Args:
            name (str): The name of the desktop to be activated.
        """
        self.ui_com_obj.ActivateDesktop(name)
        self.log.info(f'Activated the desktop with the given name({name}.')
    
    def open_baudrate_dialog(self) -> None:
        """Configures the bus parameters.
        """
        self.ui_com_obj.OpenBaudrateDialog()
        self.log.info(f'baudrate dialog opened. Configure the bus parameters.')

    def get_write_window_text_content(self) -> str:
        write_obj = Write(self)
        return write_obj.text
    
    def clear_write_window_content(self) -> None:
        write_obj = Write(self)
        write_obj.clear()
    
    def copy_write_window_content_to_clipboard(self) -> None:
        write_obj = Write(self)
        write_obj.copy()
    
    def disable_write_window_logging(self, tab_index=None) -> None:
        write_obj = Write(self)
        write_obj.disable_output_file(tab_index)
    
    def enable_write_window_logging(self, output_file: str, tab_index=None) -> None:
        write_obj = Write(self)
        write_obj.enable_output_file(output_file, tab_index)
    
    def send_text_to_write_window(self, text: str) -> None:
        write_obj = Write(self)
        write_obj.output(text)

class Write:
    """The Write object represents the Write Window in CANoe.
    It is part of the user interface.
    """
    def __init__(self, ui_obj) -> None:
        self.ui_obj = ui_obj
        self.log = self.ui_obj.log
        self.write_com_obj = win32com.client.Dispatch(self.ui_obj.ui_com_obj.Write)
    
    @property
    def text(self) -> str:
        """Gets the text contents of the Write window.

        Returns:
            str: The text content
        """
        return self.write_com_obj.Text
    
    def clear(self) -> None:
        """Clears the contents of the Write Window
        """
        self.write_com_obj.Clear()
        self.log.info(f'Cleared the contents of the Write Window.')

    def copy(self) -> None:
        """Copies the contents of the Write Window to the clipboard.
        """
        self.write_com_obj.Copy()
        self.log.info(f'Copied the contents of the Write Window to the clipboard.')

    def disable_output_file(self, tab_index=None) -> None:
        """Disables logging of all outputs of the Write Window for the certain page.

        Args:
            tab_index (int, optional): The index of the tab, for which logging of the output is to be deactivated. Defaults to None.
        """
        if tab_index:
            self.write_com_obj.DisableOutputFile(tab_index)
        else:
            self.write_com_obj.DisableOutputFile()
        self.log.info(f'Disabled logging of outputs of the Write Window. tab_index={tab_index}')

    def enable_output_file(self, output_file: str, tab_index=None) -> None:
        """Enables logging of all outputs of the Write Window in the output file for the certain page.

        Args:
            output_file (str, optional): The complete path of the output file. Defaults to None.
            tab_index (int, optional): The index of the page, for which logging of the output is to be activated. Defaults to None.
        """
        if tab_index:
            self.write_com_obj.EnableOutputFile(output_file, tab_index)
        else:
            self.write_com_obj.EnableOutputFile(output_file)
        self.log.info(f'Enabled logging of outputs of the Write Window. output_file={output_file} and tab_index={tab_index}')

    def output(self, text: str) -> None:
        """Outputs a line of text in the Write Window.

        Args:
            text (str): The text
        """
        self.write_com_obj.Output(text)
        self.log.info(f'Outputed {text} in the Write Window.')
