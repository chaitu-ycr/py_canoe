import pythoncom
import win32com.client
import win32com.client.gencache
from typing import Union
from py_canoe.utils.common import logger
from py_canoe.utils import application
from py_canoe.utils import bus as bus_utils
from py_canoe.utils import capl
from py_canoe.utils import configuration
from py_canoe.utils import measurement
from py_canoe.utils import networks
from py_canoe.utils import system
from py_canoe.utils import ui
from py_canoe.utils import version


class CANoe:
    """
    Represents a CANoe instance.
    Args:
        py_canoe_log_dir (str): The path for the CANoe log file. Defaults to an empty string.
        user_capl_functions (tuple): A tuple of user-defined CAPL function names. Defaults to an empty tuple.
    """
    def __init__(self, py_canoe_log_dir='', user_capl_functions=tuple()):
        self.bus_type = {'CAN': 1, 'J1939': 2, 'TTP': 4, 'LIN': 5, 'MOST': 6, 'Kline': 14}
        self.user_capl_functions = user_capl_functions
        self.application = application
        self.bus_utils = bus_utils
        self.capl = capl
        self.configuration = configuration
        self.measurement = measurement
        self.networks = networks
        self.system = system
        self.ui = ui
        self.version = version
        pythoncom.CoInitialize()
        self.capl_function_objects = lambda: self.measurement.MeasurementEvents.CAPL_FUNCTION_OBJECTS
        self.measurement.MeasurementEvents.CAPL_FUNCTION_NAMES = self.user_capl_functions
        self.com_object = win32com.client.gencache.EnsureDispatch("CANoe.Application")
        win32com.client.WithEvents(self.com_object, self.application.ApplicationEvents)
        win32com.client.WithEvents(self.com_object.Measurement, self.measurement.MeasurementEvents)
        win32com.client.WithEvents(self.com_object.Configuration, self.configuration.ConfigurationEvents)

    def __del__(self):
        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            logger.error(f"Error during COM uninitialization: {e}")

    def new(self, auto_save: bool = False, prompt_user: bool = False, timeout: int = 5) -> bool:
        """
        Creates a new configuration.

        Args:
            auto_save (bool): Whether to automatically save the configuration. Defaults to False.
            prompt_user (bool): Whether to prompt the user for confirmation. Defaults to False.
            timeout (int): The timeout in seconds for the operation. Defaults to 5.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.new(self, auto_save, prompt_user, timeout)

    def open(self, canoe_cfg: str, visible=True, auto_save=True, prompt_user=False, auto_stop=True, timeout=30) -> bool:
        """
        Loads a configuration.

        Args:
            canoe_cfg (str): The path to the CANoe configuration file.
            visible (bool): Whether to make the CANoe application visible. Defaults to True.
            auto_save (bool): Whether to automatically save the configuration. Defaults to True.
            prompt_user (bool): Whether to prompt the user for confirmation. Defaults to False.
            auto_stop (bool): Whether to automatically stop the measurement. Defaults to True.
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.open(self, canoe_cfg, visible, auto_save, prompt_user, auto_stop, timeout)

    def quit(self, timeout=30) -> bool:
        """
        Quits the application.

        Args:
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.quit(self, timeout)

    def get_running_instance(self, visible=True) -> Union[win32com.client.CDispatch, None]:
        """
        Gets the running instance of the CANoe application.

        Args:
            visible (bool): Whether to return only visible instances. Defaults to True.

        Returns:
            Union[win32com.client.CDispatch, None]: The running instance of the CANoe application, or None if not found.
        """
        return self.application.get_running_instance(self, visible)

    def start_measurement(self, timeout=30) -> bool:
        """
        Starts the measurement.

        Args:
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.measurement.start_measurement(self, timeout)

    def stop_measurement(self, timeout=30) -> bool:
        """
        Stops the measurement.

        Args:
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.stop_ex_measurement(timeout)

    def stop_ex_measurement(self, timeout=30) -> bool:
        """
        Stops the measurement.

        Args:
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.measurement.stop_ex_measurement(self, timeout)

    def reset_measurement(self, timeout=30) -> bool:
        """
        Resets the measurement.

        Args:
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.measurement.reset_measurement(self, timeout)

    def get_measurement_running_status(self) -> bool:
        """
        Gets the running status of the measurement.

        Returns:
            bool: True if the measurement is running, False otherwise.
        """
        return self.com_object.Measurement.Running

    def add_offline_source_log_file(self, absolute_log_file_path: str) -> bool:
        """
        Adds an offline source log file to the configuration.

        Args:
            absolute_log_file_path (str): The absolute path to the log file.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.configuration.add_offline_source_log_file(self, absolute_log_file_path)

    def start_measurement_in_animation_mode(self, animation_delay=100, timeout=30) -> bool:
        """
        Starts the measurement in animation mode.

        Args:
            animation_delay (int): The delay in milliseconds for the animation. Defaults to 100.
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.measurement.start_measurement_in_animation_mode(self, animation_delay, timeout)

    def break_measurement_in_offline_mode(self) -> bool:
        """
        Breaks the measurement in offline mode.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.measurement.break_measurement_in_offline_mode(self)

    def reset_measurement_in_offline_mode(self) -> bool:
        """
        Resets the measurement in offline mode.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.measurement.reset_measurement_in_offline_mode(self)

    def step_measurement_event_in_single_step(self) -> bool:
        """
        Steps the measurement event in single step mode.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.measurement.step_measurement_event_in_single_step(self)

    def get_measurement_index(self) -> int:
        """
        Gets the measurement index.

        Returns:
            int: The measurement index.
        """
        return self.measurement.get_measurement_index(self)

    def set_measurement_index(self, index: int) -> bool:
        """
        Sets the measurement index.

        Args:
            index (int): The measurement index to set.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.measurement.set_measurement_index(self, index)

    def save_configuration(self) -> bool:
        """
        Saves the current configuration.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.configuration.save_configuration(self)

    def save_configuration_as(self, path: str, major: int, minor: int, prompt_user: bool = False, create_dir: bool = True) -> bool:
        """
        Saves the current configuration as a new file.

        Args:
            path (str): The path to save the configuration file.
            major (int): The major version number.
            minor (int): The minor version number.
            prompt_user (bool): Whether to prompt the user for confirmation. Defaults to False.
            create_dir (bool): Whether to create the directory if it doesn't exist. Defaults to True.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.configuration.save_configuration_as(self, path, major, minor, prompt_user, create_dir)

    def get_can_bus_statistics(self, channel: int) -> dict:
        """
        Gets the CAN bus statistics.

        Args:
            channel (int): The channel number.

        Returns:
            dict: The CAN bus statistics.
        """
        return self.configuration.get_can_bus_statistics(self, channel)

    def get_canoe_version_info(self) -> dict:
        """
        Gets the version information of the CANoe application.

        Returns:
            dict: The version information.
        """
        return self.version.get_canoe_version_info(self)

    def get_bus_databases_info(self, bus: str = 'CAN') -> dict:
        """
        Gets the bus databases information.

        Returns:
            dict: The bus databases information.
        """
        return self.bus_utils.get_bus_databases_info(self, bus)

    def get_bus_nodes_info(self, bus: str = 'CAN') -> dict:
        """
        Gets the bus nodes information.

        Returns:
            dict: The bus nodes information.
        """
        return self.bus_utils.get_bus_nodes_info(self, bus)

    def get_signal_value(self, bus: str, channel: int, message: str, signal: str, raw_value: bool = False) -> Union[int, float, None]:
        """
        Gets the value of a signal.

        Args:
            bus (str): The bus name.
            channel (int): The channel number.
            message (str): The message name.
            signal (str): The signal name.
            raw_value (bool): Whether to get the raw value. Defaults to False.

        Returns:
            Union[int, float, None]: The signal value or None if not found.
        """
        return self.bus_utils.get_signal_value(self, bus, channel, message, signal, raw_value)

    def set_signal_value(self, bus: str, channel: int, message: str, signal: str, value: Union[int, float], raw_value: bool = False) -> bool:
        """
        Sets the value of a signal.

        Args:
            bus (str): The bus name.
            channel (int): The channel number.
            message (str): The message name.
            signal (str): The signal name.
            value (Union[int, float]): The value to set.
            raw_value (bool): Whether to set the raw value. Defaults to False.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.bus_utils.set_signal_value(self, bus, channel, message, signal, value, raw_value)

    def get_signal_full_name(self, bus: str, channel: int, message: str, signal: str) -> Union[str, None]:
        """
        Gets the full name of a signal.

        Args:
            bus (str): The bus name.
            channel (int): The channel number.
            message (str): The message name.
            signal (str): The signal name.

        Returns:
            Union[str, None]: The full name of the signal or None if not found.
        """
        return self.bus_utils.get_signal_full_name(self, bus, channel, message, signal)

    def check_signal_online(self, bus: str, channel: int, message: str, signal: str) -> bool:
        """
        Checks if a signal is online.

        Args:
            bus (str): The bus name.
            channel (int): The channel number.
            message (str): The message name.
            signal (str): The signal name.

        Returns:
            bool: True if the signal is online, False otherwise.
        """
        return self.bus_utils.check_signal_online(self, bus, channel, message, signal)

    def check_signal_state(self, bus: str, channel: int, message: str, signal: str) -> int:
        """
        Checks the state of a signal.

        Args:
            bus (str): The bus name.
            channel (int): The channel number.
            message (str): The message name.
            signal (str): The signal name.

        Returns:
            int: The state of the signal.
        """
        return self.bus_utils.check_signal_state(self, bus, channel, message, signal)

    def get_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, raw_value=False) -> Union[float, int, None]:
        """
        Gets the value of a J1939 signal.

        Args:
            bus (str): The bus name.
            channel (int): The channel number.
            message (str): The message name.
            signal (str): The signal name.
            source_addr (int): The source address.
            dest_addr (int): The destination address.
            raw_value (bool): Whether to get the raw value. Defaults to False.

        Returns:
            Union[float, int, None]: The signal value or None if not found.
        """
        return self.bus_utils.get_j1939_signal_value(self, bus, channel, message, signal, source_addr, dest_addr, raw_value)

    def set_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, value: Union[float, int], raw_value: bool = False) -> bool:
        """
        Sets the value of a J1939 signal.

        Args:
            bus (str): The bus name.
            channel (int): The channel number.
            message (str): The message name.
            signal (str): The signal name.
            source_addr (int): The source address.
            dest_addr (int): The destination address.
            value (Union[float, int]): The value to set.
            raw_value (bool): Whether to set the raw value. Defaults to False.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.bus_utils.set_j1939_signal_value(self, bus, channel, message, signal, source_addr, dest_addr, value, raw_value)

    def get_j1939_signal_full_name(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> Union[str, None]:
        """
        Gets the full name of a J1939 signal.

        Args:
            bus (str): The bus name.
            channel (int): The channel number.
            message (str): The message name.
            signal (str): The signal name.
            source_addr (int): The source address.
            dest_addr (int): The destination address.

        Returns:
            Union[str, None]: The full name of the signal or None if not found.
        """
        return self.bus_utils.get_j1939_signal_full_name(self, bus, channel, message, signal, source_addr, dest_addr)

    def check_j1939_signal_online(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> bool:
        """
        Checks if a J1939 signal is online.

        Args:
            bus (str): The bus name.
            channel (int): The channel number.
            message (str): The message name.
            signal (str): The signal name.
            source_addr (int): The source address.
            dest_addr (int): The destination address.

        Returns:
            bool: True if the signal is online, False otherwise.
        """
        return self.bus_utils.check_j1939_signal_online(self, bus, channel, message, signal, source_addr, dest_addr)

    def check_j1939_signal_state(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> int:
        """
        Checks the state of a J1939 signal.

        Args:
            bus (str): The bus name.
            channel (int): The channel number.
            message (str): The message name.
            signal (str): The signal name.
            source_addr (int): The source address.
            dest_addr (int): The destination address.

        Returns:
            int: The state of the signal.
        """
        return self.bus_utils.check_j1939_signal_state(self, bus, channel, message, signal, source_addr, dest_addr)

    def ui_activate_desktop(self, name: str) -> bool:
        """
        Activates a desktop by name.

        Args:
            name (str): The name of the desktop to activate.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.ui.activate_desktop(self, name)

    def ui_open_baudrate_dialog(self) -> bool:
        """
        Opens the baudrate dialog.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.ui.open_baudrate_dialog(self)

    def write_text_in_write_window(self, text: str) -> bool:
        """
        Writes text in the write window.

        Args:
            text (str): The text to write.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.ui.write_text_in_write_window(self, text)

    def read_text_from_write_window(self) -> Union[str, None]:
        """
        Reads text from the write window.

        Returns:
            Union[str, None]: The text from the write window or None if not found.
        """
        return self.ui.read_text_from_write_window(self)

    def clear_write_window_content(self) -> bool:
        """
        Clears the content of the write window.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.ui.clear_write_window_content(self)

    def copy_write_window_content(self) -> bool:
        """
        Copies the content of the write window.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.ui.copy_write_window_content(self)

    def enable_write_window_output_file(self, output_file: str, tab_index=None) -> bool:
        """
        Enables the write window output file.

        Args:
            output_file (str): The output file path.
            tab_index (Optional[int]): The tab index to enable the output file for.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.ui.enable_write_window_output_file(self, output_file, tab_index)

    def disable_write_window_output_file(self, tab_index=None) -> bool:
        """
        Disables the write window output file.

        Args:
            tab_index (Optional[int]): The tab index to disable the output file for.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.ui.disable_write_window_output_file(self, tab_index)

    def define_system_variable(self, sys_var_name: str, value: Union[int, float, str], read_only: bool = False) -> object:
        """
        Defines a system variable.

        Args:
            sys_var_name (str): The name of the system variable.
            value (Union[int, float, str]): The value of the system variable.
            read_only (bool): Whether the system variable is read-only.

        Returns:
            object: The created system variable object.
        """
        return self.system.add_system_variable(self, sys_var_name, value, read_only)

    def get_system_variable_value(self, sys_var_name: str, return_symbolic_name=False) -> Union[int, float, str, None]:
        """
        Gets the value of a system variable.

        Args:
            sys_var_name (str): The name of the system variable.
            return_symbolic_name (bool): Whether to return the symbolic name.

        Returns:
            Union[int, float, str, None]: The value of the system variable or None if not found.
        """
        return self.system.get_system_variable_value(self, sys_var_name, return_symbolic_name)

    def set_system_variable_value(self, sys_var_name: str, value: Union[int, float, str]) -> bool:
        """
        Sets the value of a system variable.

        Args:
            sys_var_name (str): The name of the system variable.
            value (Union[int, float, str]): The value to set.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.system.set_system_variable_value(self, sys_var_name, value)

    def set_system_variable_array_values(self, sys_var_name: str, value: tuple, index: int = 0) -> bool:
        """
        Sets the values of a system variable array.

        Args:
            sys_var_name (str): The name of the system variable.
            value (tuple): The values to set.
            index (int): The index to set the values at.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.system.set_system_variable_array_values(self, sys_var_name, value, index)

    def _fetch_diagnostic_devices(self):
        """
        Fetches the diagnostic devices.
        """
        return self.networks.fetch_diagnostic_devices(self)

    def send_diag_request(self, diag_ecu_qualifier_name: str, request: str, request_in_bytes=True, return_sender_name=False, response_in_bytearray=False) -> Union[str, dict]:
        """
        Sends a diagnostic request.

        Args:
            diag_ecu_qualifier_name (str): The diagnostic ECU qualifier name.
            request (str): The diagnostic request.
            request_in_bytes (bool): Whether the request is in bytes.
            return_sender_name (bool): Whether to return the sender name.
            response_in_bytearray (bool): Whether to return the response in bytearray.

        Returns:
            Union[str, dict]: The response from the diagnostic request.
        """
        return self.networks.send_diag_request(self, diag_ecu_qualifier_name, request, request_in_bytes, return_sender_name, response_in_bytearray)

    def control_tester_present(self, diag_ecu_qualifier_name: str, value: bool) -> bool:
        """
        Controls the tester present signal.

        Args:
            diag_ecu_qualifier_name (str): The diagnostic ECU qualifier name.
            value (bool): The value to set for the tester present signal.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.networks.control_tester_present(self, diag_ecu_qualifier_name, value)

    def set_replay_block_file(self, block_name: str, recording_file_path: str) -> bool:
        """
        Sets the replay block file.

        Args:
            block_name (str): The name of the replay block.
            recording_file_path (str): The path to the recording file.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.configuration.set_replay_block_file(self, block_name, recording_file_path)

    def control_replay_block(self, block_name: str, start_stop: bool) -> bool:
        """
        Controls the replay block.

        Args:
            block_name (str): The name of the replay block.
            start_stop (bool): True to start the replay block, False to stop it.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.configuration.control_replay_block(self, block_name, start_stop)

    def enable_disable_replay_block(self, block_name: str, enable_disable: bool) -> bool:
        """
        Enables or disables a replay block.

        Args:
            block_name (str): The name of the replay block.
            enable_disable (bool): True to enable the replay block, False to disable it.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.configuration.enable_disable_replay_block(self, block_name, enable_disable)

    def compile_all_capl_nodes(self, wait_time: Union[int, float] = 5) -> Union[capl.CompileResult, None]:
        """
        Compiles all CAPL nodes in the application.

        Args:
            wait_time (Union[int, float]): The time to wait for the compilation to complete.

        Returns:
            Union[capl.CompileResult, None]: The compilation result or None if an error occurred.
        """
        return self.capl.compile_all_capl_nodes(self, wait_time)

    def call_capl_function(self, name: str, *arguments) -> bool:
        """
        Calls a CAPL function.

        Args:
            name (str): The name of the CAPL function.
            *arguments: The arguments to pass to the CAPL function.

        Returns:
            bool: True if the function call was successful, False otherwise.
        """
        return self.capl.call_capl_function(self, name, *arguments)
