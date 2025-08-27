
from typing import TYPE_CHECKING, Iterable
if TYPE_CHECKING:
    from py_canoe.core.conf_children.measurement_setup import Logging, ExporterSymbol, Message

import gc
import pythoncom
from typing import Union

from py_canoe.core.application import Application
from py_canoe.core.capl import CompileResult
from py_canoe.utils.common import logger, wait


class CANoe:
    def __init__(self, py_canoe_log_dir='', user_capl_functions=tuple()):
        try:
            pythoncom.CoInitialize()
        except pythoncom.com_error:
            logger.warning("⚠️ COM already initialized in this thread.")
        self.user_capl_functions = user_capl_functions
        self.application: Application = None

    def __del__(self):
        try:
            wait(0.5)
            pythoncom.CoUninitialize()
            wait(0.5)
        except Exception as e:
            logger.error(f"❌ Error during COM uninitialization: {e}")

    def _reset_application(self):
        try:
            wait(0.5)
            if self.application:
                del self.application
                self.application = None
            wait(0.5)
            gc.collect()
            wait(0.5)
        except Exception as e:
            logger.error(f"❌ Error during application reset: {e}")

    def new(self, auto_save=False, prompt_user=False, timeout=5) -> bool:
        """
        Creates a new configuration.

        Args:
            auto_save (bool): Whether to automatically save the configuration. Defaults to False.
            prompt_user (bool): Whether to prompt the user for confirmation. Defaults to False.
            timeout (int): The timeout in seconds for the operation. Defaults to 5.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        self._reset_application()
        self.application = Application()
        return self.application.new(auto_save, prompt_user, timeout)

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
        self._reset_application()
        self.application = Application()
        self.application.user_capl_functions = self.user_capl_functions
        return self.application.open(canoe_cfg, visible, auto_save, prompt_user, timeout)

    def quit(self, timeout=30) -> bool:
        """
        Quits the application.

        Args:
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        status = self.application.quit(timeout)
        self._reset_application()
        return status

    def attach_to_active_application(self) -> bool:
        """
        Attach to a active instance of the CANoe application.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        self._reset_application()
        self.application = Application()
        self.application.user_capl_functions = self.user_capl_functions
        return self.application.attach_to_active_application()

    def save_configuration(self) -> bool:
        """
        Saves the current configuration.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.configuration.save()

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
        return self.application.configuration.save_as(path, major, minor, prompt_user, create_dir)

    def start_measurement(self, timeout=30) -> bool:
        """
        Starts the measurement.

        Args:
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.measurement.start(timeout)

    def stop_measurement(self, timeout=30) -> bool:
        """
        Stops the measurement.

        Args:
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.measurement.stop(timeout)

    def stop_ex_measurement(self, timeout=30) -> bool:
        """
        Stops the measurement.

        Args:
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.measurement.stop_ex(timeout)

    def reset_measurement(self, timeout=30) -> bool:
        """
        Restarts the measurement if running.

        Args:
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        if self.application.measurement.running:
            stop_status = self.stop_measurement(timeout)
            start_status = self.start_measurement(timeout)
            return stop_status and start_status
        else:
            logger.warning("⚠️ Measurement is not running, cannot reset.")
            return False

    def get_measurement_running_status(self) -> bool:
        """
        Gets the running status of the measurement.

        Returns:
            bool: True if the measurement is running, False otherwise.
        """
        return self.application.measurement.running

    def add_offline_source_log_file(self, absolute_log_file_path: str) -> bool:
        """
        Adds an offline source log file to the configuration.

        Args:
            absolute_log_file_path (str): The absolute path to the log file.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.configuration.add_offline_source_log_file(absolute_log_file_path)

    def start_measurement_in_animation_mode(self, animation_delay=100, timeout=30) -> bool:
        """
        Starts the measurement in animation mode.

        Args:
            animation_delay (int): The delay in milliseconds for the animation. Defaults to 100.
            timeout (int): The timeout in seconds for the operation. Defaults to 30.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.measurement.start_measurement_in_animation_mode(animation_delay, timeout)

    def break_measurement_in_offline_mode(self) -> bool:
        """
        Breaks the measurement in offline mode.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.measurement.break_measurement_in_offline_mode()

    def reset_measurement_in_offline_mode(self) -> bool:
        """
        Resets the measurement in offline mode.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.measurement.reset_measurement_in_offline_mode()

    def step_measurement_event_in_single_step(self) -> bool:
        """
        Steps the measurement event in single step mode.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.measurement.process_measurement_event_in_single_step()

    def get_measurement_index(self) -> int:
        """
        Gets the measurement index.

        Returns:
            int: The measurement index.
        """
        return self.application.measurement.measurement_index

    def set_measurement_index(self, index: int) -> bool:
        """
        Sets the measurement index.

        Args:
            index (int): The measurement index to set.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        self.application.measurement.measurement_index = index
        return True

    def get_can_bus_statistics(self, channel: int) -> dict:
        """
        Gets the CAN bus statistics.

        Args:
            channel (int): The channel number.

        Returns:
            dict: The CAN bus statistics.
        """
        return self.application.configuration.get_can_bus_statistics(channel)

    def get_canoe_version_info(self) -> dict:
        """
        Gets the version information of the CANoe application.

        Returns:
            dict: The version information.
        """
        return self.application.version.get_canoe_version_info()

    def get_bus_databases_info(self, bus: str = 'CAN') -> dict:
        """
        Gets the bus databases information.

        Returns:
            dict: The bus databases information.
        """
        return self.application.bus.get_bus_databases_info(bus)

    def get_bus_nodes_info(self, bus: str = 'CAN') -> dict:
        """
        Gets the bus nodes information.

        Returns:
            dict: The bus nodes information.
        """
        return self.application.bus.get_bus_nodes_info(bus)

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
        return self.application.bus.get_signal_value(bus, channel, message, signal, raw_value)

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
        return self.application.bus.set_signal_value(bus, channel, message, signal, value, raw_value)

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
        return self.application.bus.get_signal_full_name(bus, channel, message, signal)

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
        return self.application.bus.check_signal_online(bus, channel, message, signal)

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
        return self.application.bus.check_signal_state(bus, channel, message, signal)

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
        return self.application.bus.get_j1939_signal_value(bus, channel, message, signal, source_addr, dest_addr, raw_value)

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
        return self.application.bus.set_j1939_signal_value(bus, channel, message, signal, source_addr, dest_addr, value, raw_value)

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
        return self.application.bus.get_j1939_signal_full_name(bus, channel, message, signal, source_addr, dest_addr)

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
        return self.application.bus.check_j1939_signal_online(bus, channel, message, signal, source_addr, dest_addr)

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
        return self.application.bus.check_j1939_signal_state(bus, channel, message, signal, source_addr, dest_addr)

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
        return self.application.system.add_system_variable(sys_var_name, value, read_only)

    def get_system_variable_value(self, sys_var_name: str, return_symbolic_name=False) -> Union[int, float, str, None]:
        """
        Gets the value of a system variable.

        Args:
            sys_var_name (str): The name of the system variable.
            return_symbolic_name (bool): Whether to return the symbolic name.

        Returns:
            Union[int, float, str, None]: The value of the system variable or None if not found.
        """
        return self.application.system.get_system_variable_value(sys_var_name, return_symbolic_name)

    def set_system_variable_value(self, sys_var_name: str, value: Union[int, float, str]) -> bool:
        """
        Sets the value of a system variable.

        Args:
            sys_var_name (str): The name of the system variable.
            value (Union[int, float, str]): The value to set.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.system.set_system_variable_value(sys_var_name, value)

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
        return self.application.system.set_system_variable_array_values(sys_var_name, value, index)

    def get_environment_variable_value(self, env_var_name: str) -> Union[int, float, str, tuple, None]:
        """
        returns a environment variable value.

        Args:
            env_var_name (str): The name of the environment variable. Ex- "float_var"

        Returns:
            Environment Variable value.
        """
        return self.application.environment.get_environment_variable_value(env_var_name)

    def set_environment_variable_value(self, env_var_name: str, value: Union[int, float, str, tuple]) -> bool:
        """
        Sets the value of an environment variable.

        Args:
            env_var_name (str): The name of the environment variable. Ex- "speed".
            value (Union[int, float, str, tuple]): variable value. supported CAPL environment variable data types integer, double, string and data.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.environment.set_environment_variable_value(env_var_name, value)

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
        return self.application.networks.send_diag_request(diag_ecu_qualifier_name, request, request_in_bytes, return_sender_name, response_in_bytearray)

    def control_tester_present(self, diag_ecu_qualifier_name: str, value: bool) -> bool:
        """
        Controls the tester present signal.

        Args:
            diag_ecu_qualifier_name (str): The diagnostic ECU qualifier name.
            value (bool): The value to set for the tester present signal.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.networks.control_tester_present(diag_ecu_qualifier_name, value)

    def set_replay_block_file(self, block_name: str, recording_file_path: str) -> bool:
        """
        Sets the replay block file.

        Args:
            block_name (str): The name of the replay block.
            recording_file_path (str): The path to the recording file.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.configuration.set_replay_block_file(block_name, recording_file_path)

    def control_replay_block(self, block_name: str, start_stop: bool) -> bool:
        """
        Controls the replay block.

        Args:
            block_name (str): The name of the replay block.
            start_stop (bool): True to start the replay block, False to stop it.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.configuration.control_replay_block(block_name, start_stop)

    def enable_disable_replay_block(self, block_name: str, enable_disable: bool) -> bool:
        """
        Enables or disables a replay block.

        Args:
            block_name (str): The name of the replay block.
            enable_disable (bool): True to enable the replay block, False to disable it.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.configuration.enable_disable_replay_block(block_name, enable_disable)

    def compile_all_capl_nodes(self, wait_time: Union[int, float] = 5) -> Union[CompileResult, None]:
        """
        Compiles all CAPL nodes in the application.

        Args:
            wait_time (Union[int, float]): The time to wait for the compilation to complete.

        Returns:
            The compilation result or None if an error occurred.
        """
        return self.application.capl.compile(wait_time)

    def call_capl_function(self, name: str, *arguments) -> bool:
        """
        Calls a CAPL function.

        Args:
            name (str): The name of the CAPL function.
            *arguments: The arguments to pass to the CAPL function.

        Returns:
            bool: True if the function call was successful, False otherwise.
        """
        return self.application.capl.call_capl_function(name, *arguments)

    def get_test_environments(self) -> dict:
        """returns dictionary of test environment names and class."""
        return self.application.configuration.get_test_environments()

    def get_test_modules(self, env_name: str) -> dict:
        """returns dictionary of test environment test module names and its class object.

        Args:
            env_name (str): test environment name. avoid duplicate test environment names in CANoe configuration.
        """
        return self.application.configuration.get_test_modules(env_name)

    def execute_test_module(self, test_module_name: str) -> int:
        """use this method to execute test module.

        Args:
            test_module_name (str): test module name. avoid duplicate test module names in CANoe configuration.

        Returns:
            int: test module execution verdict. 0 ='VerdictNotAvailable', 1 = 'VerdictPassed', 2 = 'VerdictFailed',
        """
        return self.application.configuration.execute_test_module(test_module_name)

    def stop_test_module(self, test_module_name: str):
        """stops execution of test module.

        Args:
            test_module_name (str): test module name. avoid duplicate test module names in CANoe configuration.
        """
        return self.application.configuration.stop_test_module(test_module_name)

    def execute_all_test_modules_in_test_env(self, env_name: str):
        """executes all test modules available in test environment.

        Args:
            env_name (str): test environment name. avoid duplicate test environment names in CANoe configuration.
        """
        return self.application.configuration.execute_all_test_modules_in_test_env(env_name)

    def stop_all_test_modules_in_test_env(self, env_name: str):
        """stops execution of all test modules available in test environment.

        Args:
            env_name (str): test environment name. avoid duplicate test environment names in CANoe configuration.
        """
        return self.application.configuration.stop_all_test_modules_in_test_env(env_name)

    def execute_all_test_environments(self):
        """executes all test environments available in test setup."""
        return self.application.configuration.execute_all_test_environments()

    def stop_all_test_environments(self):
        """stops execution of all test environments available in test setup."""
        return self.application.configuration.stop_all_test_environments()

    def ui_activate_desktop(self, name: str) -> bool:
        """
        Activates a desktop by name.

        Args:
            name (str): The name of the desktop to activate.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.ui.activate_desktop(name)

    def ui_open_baudrate_dialog(self) -> bool:
        """
        Opens the baudrate dialog.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.ui.open_baudrate_dialog()

    def write_text_in_write_window(self, text: str) -> bool:
        """
        Writes text in the write window.

        Args:
            text (str): The text to write.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.ui.write.output(text)

    def read_text_from_write_window(self) -> Union[str, None]:
        """
        Reads text from the write window.

        Returns:
            Union[str, None]: The text from the write window or None if not found.
        """
        return self.application.ui.write.text

    def clear_write_window_content(self) -> bool:
        """
        Clears the content of the write window.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.ui.write.clear()

    def copy_write_window_content(self) -> bool:
        """
        Copies the content of the write window.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.ui.write.copy()

    def enable_write_window_output_file(self, output_file: str, tab_index=None) -> bool:
        """
        Enables the write window output file.

        Args:
            output_file (str): The output file path.
            tab_index (Optional[int]): The tab index to enable the output file for.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.ui.write.enable_output_file(output_file, tab_index)

    def disable_write_window_output_file(self, tab_index=None) -> bool:
        """
        Disables the write window output file.

        Args:
            tab_index (Optional[int]): The tab index to disable the output file for.

        Returns:
            bool: True if the operation was successful, False otherwise.
        """
        return self.application.ui.write.disable_output_file(tab_index)

    def add_database(self, database_file: str, database_network: str, database_channel: int) -> bool:
        """adds database file to a network channel

        Args:
            database_file (str): database file to attach. give full file path.
            database_network (str): network name on which you want to add this database.
            database_channel (int): channel name on which you want to add this database.
        """
        return self.application.configuration.add_database(database_file, database_network, database_channel)

    def remove_database(self, database_file: str, database_channel: int) -> bool:
        """remove database file from a channel

        Args:
            database_file (str): database file to remove. give full file path.
            database_channel (int): channel name on which you want to remove database.
        """
        return self.application.configuration.remove_database(database_file, database_channel)

    def get_logging_blocks(self) -> list['Logging']:
        """Return all available logging blocks."""
        return list(self.application.configuration.get_logging_blocks())

    def add_logging_block(self, full_name: str) -> 'Logging':
        """adds a new logging block to configuration measurement setup.

        Args:
            full_name (str): full path to log file as "C:/file.(asc|blf|mf4|...)", may have field functions like {IncMeasurement} in the file name.

        Returns:
            Logging: returns Logging object of added logging block.
        """
        return self.application.configuration.add_logging_block(full_name)

    def remove_logging_block(self, index: int) -> None:
        """removes a logging block from configuration measurement setup.

        Args:
            index (int): index of logging block to remove. logging blocks indexing starts from 1 and not 0.
        """
        return self.application.configuration.remove_logging_block(index)

    def load_logs_for_exporter(self, logger_index: int) -> None:
        """Load all source files of exporter and determine symbols/messages.

        Args:
            logger_index (int): indicates logger and its log files
        """
        return self.application.configuration.load_logs_for_exporter(logger_index)

    def get_symbols(self, logger_index: int) -> list['ExporterSymbol']:
        """Return all exporter symbols from given logger."""
        return self.application.configuration.get_symbols(logger_index)

    def get_messages(self, logger_index: int) -> list['Message']:
        """Return all messages from given logger."""
        return self.application.configuration.get_messages(logger_index)

    def add_filters_to_exporter(self, logger_index: int, full_names: 'Iterable'):
        """Add messages and symbols to exporter filter by their full names.

        Args:
            logger_index (int): indicates logger
            full_names (Iterable): full names of messages and symbols
        """
        return self.application.configuration.add_filters_to_exporter(logger_index, full_names)

    def start_export(self, logger_index: int):
        """Starts the export/conversion of exporter.

        Args:
            logger_index (int): indicates logger
        """
        return self.application.configuration.start_export(logger_index)

    def set_configuration_modified(self, modified: bool) -> None:
        """Change status of configuration.

        Args:
            modified (bool): True if configuration is modified, False otherwise.
        """
        return self.application.configuration.set_configuration_modified(modified)
