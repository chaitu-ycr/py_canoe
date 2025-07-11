import logging
import win32com.client

logging.getLogger('py_canoe')

class Ui:
    def __init__(self, app):
        self.app = app
        self.com_object = win32com.client.Dispatch(self.app.com_object.UI)

    def get_command_enabled(self, command: str) -> bool:
        return self.com_object.GetCommandEnabled(command)

    def set_command_enabled(self, command: str, enabled: bool) -> None:
        self.com_object.SetCommandEnabled(command, enabled)

    @property
    def write(self) -> 'Write':
        return Write(self)

    def activate_desktop(self, desktop_name: str) -> None:
        self.com_object.ActivateDesktop(desktop_name)

    def create_desktop(self, desktop_name: str) -> None:
        if float(f"{self.app.version.major}.{self.app.version.minor}") >= 15.3:
            self.com_object.CreateDesktop(desktop_name)

    def open_baudrate_dialog(self) -> None:
        self.com_object.OpenBaudrateDialog()

class Write:
    def __init__(self, ui: Ui):
        self.com_object = win32com.client.Dispatch(ui.com_object.Write)

    @property
    def text(self) -> str:
        return self.com_object.Text

    def clear(self) -> None:
        self.com_object.Clear()

    def copy(self) -> None:
        self.com_object.Copy()

    def disable_output_file(self, tab_index: int = 0) -> None:
        self.com_object.DisableOutputFile(tab_index)

    def enable_output_file(self, output_file: str, tab_index: int = 0) -> None:
        self.com_object.EnableOutputFile(output_file, tab_index)

    def output(self, text: str) -> None:
        self.com_object.Output(text)