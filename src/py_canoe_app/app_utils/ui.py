# import external modules here
import logging
import win32com.client


class Ui:
    """The UI object represents the user interface in CANoe."""
    def __init__(self, app_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.UI)
            self.write_window_com_obj = win32com.client.Dispatch(self.com_obj.Write)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe UI: {str(e)}')

    def get_command_availability(self, command: str) -> bool:
        return self.com_obj.CommandEnabled(command)

    def set_command_availability(self, command: str, value: bool) -> None:
        ce_obj = self.com_obj.CommandEnabled(command)
        ce_obj = value

    def activate_desktop(self, name: str) -> None:
        self.com_obj.ActivateDesktop(name)

    def open_baudrate_dialog(self) -> None:
        self.com_obj.OpenBaudrateDialog()

    @property
    def get_write_window_text(self) -> str:
        return self.write_window_com_obj.Text

    def clear_write_window(self) -> None:
        self.write_window_com_obj.Clear()

    def copy_write_window_content(self) -> None:
        self.write_window_com_obj.Copy()

    def disable_write_window_output_file(self, tab_index=None) -> None:
        if tab_index:
            self.write_window_com_obj.DisableOutputFile(tab_index)
        else:
            self.write_window_com_obj.DisableOutputFile()

    def enable_write_window_output_file(self, output_file: str, tab_index=None) -> None:
        if tab_index:
            self.write_window_com_obj.EnableOutputFile(output_file, tab_index)
        else:
            self.write_window_com_obj.EnableOutputFile(output_file)

    def output_text_in_write_window(self, text: str) -> None:
        self.write_window_com_obj.Output(text)
