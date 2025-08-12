from typing import Union

from py_canoe.utils.common import logger
from py_canoe.core.version import Version


class Ui:
    """
    The UI object represents the user interface in CANoe.
    """
    def __init__(self, app):
        self.app = app
        self.com_object = self.app.com_object.UI

    def get_command_enabled(self, command: str) -> bool:
        return self.com_object.GetCommandEnabled(command)

    def set_command_enabled(self, command: str, enabled: bool) -> None:
        self.com_object.SetCommandEnabled(command, enabled)

    @property
    def write(self) -> 'Write':
        return Write(self.com_object.Write)

    def activate_desktop(self, desktop_name: str) -> None:
        self.com_object.ActivateDesktop(desktop_name)

    def create_desktop(self, desktop_name: str) -> bool:
        version = Version(self.app)
        if float(f"{version.major}.{version.minor}") >= 15.3:
            self.com_object.CreateDesktop(desktop_name)
            return True
        else:
            logger.warning(f"😡 Cannot create desktop '{desktop_name}': Requires CANoe version 15.3 or higher.")
            return False

    def open_baudrate_dialog(self) -> None:
        self.com_object.OpenBaudrateDialog()


class Write:
    def __init__(self, write):
        self.com_object = write

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


def activate_desktop(app, name: str) -> bool:
    try:
        ui = Ui(app)
        ui.activate_desktop(name)
        logger.info(f"📢 UI Desktop '{name}' activated successfully")
        return True
    except Exception as e:
        logger.error(f"😡 Error activating UI Desktop '{name}': {e}")
        return False

def create_desktop(app, name: str) -> bool:
    try:
        ui = Ui(app)
        status = ui.create_desktop(name)
        if status:
            logger.info(f"📢 UI Desktop '{name}' created successfully")
        return status
    except Exception as e:
        logger.error(f"😡 Error creating UI Desktop '{name}': {e}")
        return False

def open_baudrate_dialog(app) -> bool:
    try:
        ui = Ui(app)
        ui.open_baudrate_dialog()
        logger.info("📢 UI Baudrate Dialog opened successfully")
        return True
    except Exception as e:
        logger.error(f"😡 Error opening UI Baudrate Dialog: {e}")
        return False

def write_text_in_write_window(app, text: str) -> bool:
    try:
        ui = Ui(app)
        ui.write.output(text)
        logger.info(f"✍️ Text written in write window: {text}")
        return True
    except Exception as e:
        logger.error(f"😡 Error writing text in write window: {e}")
        return False

def read_text_from_write_window(app) -> Union[str, None]:
    try:
        ui = Ui(app)
        text = ui.write.text
        logger.info("📖 Text read successfully from write window")
        for line in text.splitlines():
            logger.info(f"    {line}")
        return text
    except Exception as e:
        logger.error(f"😡 Error reading text from write window: {e}")
        return None

def clear_write_window_content(app) -> bool:
    try:
        ui = Ui(app)
        ui.write.clear()
        logger.info("🧹 Write Window content cleared successfully")
        return True
    except Exception as e:
        logger.error(f"😡 Error clearing write window content: {e}")
        return False

def copy_write_window_content(app) -> bool:
    try:
        ui = Ui(app)
        ui.write.copy()
        logger.info("📷 Write Window content copied to clipboard successfully")
        return True
    except Exception as e:
        logger.error(f"😡 Error copying write window content: {e}")
        return False

def enable_write_window_output_file(app, output_file: str, tab_index=None) -> bool:
    try:
        ui = Ui(app)
        if tab_index is not None:
            ui.write.enable_output_file(output_file, tab_index)
        else:
            ui.write.enable_output_file(output_file)
        logger.info(f"✔️ Enabled write window output file: {output_file}")
        return True
    except Exception as e:
        logger.error(f"😡 Error enabling write window output file: {e}")
        return False

def disable_write_window_output_file(app, tab_index=None) -> bool:
    try:
        ui = Ui(app)
        if tab_index is not None:
            ui.write.disable_output_file(tab_index)
        else:
            ui.write.disable_output_file()
        logger.info("⏹️ Disabled write window output file")
        return True
    except Exception as e:
        logger.error(f"😡 Error disabling write window output file: {e}")
        return False
