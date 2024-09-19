# import external modules here
import os
import sys
from typing import Union
from time import sleep as wait

# import internal modules here
from py_canoe_app.py_canoe_logger import PyCanoeLogger
from py_canoe_app.application import Application


class CANoe:
    """
    Represents a CANoe instance.
    Args:
        py_canoe_log_dir (str): The path for the CANoe log file. Defaults to an empty string.
        user_capl_functions (tuple): A tuple of user-defined CAPL function names. Defaults to an empty tuple.
    """

    def __init__(self, py_canoe_log_dir='', user_capl_functions=tuple()):
        try:
            self.__log = PyCanoeLogger(py_canoe_log_dir).log
            self.application = Application(user_capl_functions)
            self.__diag_devices = dict()
            self.__test_environments = dict()
            self.__test_modules = list()
            self.__replay_blocks = dict()
            self.__namespaces = dict()
            self.__variable_files = dict()
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe: {str(e)}')
            sys.exit(1)

    def new(self, auto_save=False, prompt_user=False) -> None:
        """Creates a new configuration.

        Args:
            auto_save (bool, optional): A boolean value that indicates whether the active configuration should be saved if it has been changed. Defaults to False.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations. Defaults to False.
        """
        try:
            self.stop_ex_measurement()
            self.application.new(auto_save, prompt_user)
            self.__log.debug('👉 created a new configuration.')
        except Exception as e:
            self.__log.error(f'😡 Error creating new configuration: {str(e)}')
            sys.exit(1)

    def open(self, canoe_cfg: str, visible=True, auto_save=False, prompt_user=False, auto_stop=False) -> None:
        """Loads CANoe configuration.

        Args:
            canoe_cfg (str): The complete path for the CANoe configuration.
            visible (bool): True if you want to see CANoe UI. Defaults to True.
            auto_save (bool, optional): A boolean value that indicates whether the active configuration should be saved if it has been changed. Defaults to False.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations. Defaults to False.
            auto_stop (bool, optional): A boolean value that indicates whether to stop the measurement before opening the configuration. Defaults to False.
        """
        self.application.visible = visible
        if self.application.measurement.running and not auto_stop:
            self.__log.error('😡 Measurement is running. Stop the measurement or set argument auto_stop=True.')
            sys.exit(1)
        elif self.application.measurement.running and auto_stop:
            self.__log.error('😇 Active Measurement is running. Stopping measurement before opening your configuration.')
            self.stop_ex_measurement()
        if not auto_save and self.application.configuration.modified:
            self.__log.error(f'😡 Active CANoe configuration has unsaved changes. Save changes or set argument auto_save=True.')
            sys.exit(1)
        if os.path.isfile(canoe_cfg):
            self.__log.debug(f'🔎 CANoe configuration "{canoe_cfg}" found.')
            self.application.open(canoe_cfg, auto_save, prompt_user)
            self.__diag_devices = self.application.networks.fetch_all_diag_devices()
            self.__test_environments = self.application.configuration.get_all_test_setup_environments()
            self.__test_modules = self.application.configuration.get_all_test_modules_in_test_environments()
            self.__namespaces = self.application.system.fetch_namespaces()
            self.__variable_files = self.application.system.fetch_variables_files()
            self.__replay_blocks = self.application.configuration.simulation_setup.replay_collection.fetch_replay_blocks()
            self.__log.debug(f'📢 loaded CANoe configuration successfully 🎉')
        else:
            self.__log.error(f'😡 CANoe configuration "{canoe_cfg}" not found.')
            sys.exit(1)

    def quit(self):
        """Quits CANoe without saving changes in the configuration."""
        try:
            self.application.quit()
            self.__log.debug('📢 CANoe Application Closed.')
        except Exception as e:
            self.__log.error(f'😡 Error quitting CANoe application: {str(e)}')
            sys.exit(1)

    def start_measurement(self, timeout=60) -> bool:
        """Starts the measurement.

        Args:
            timeout (int, optional): measurement start/stop event timeout in seconds. Defaults to 60.

        Returns:
            True if measurement started. else False.
        """
        try:
            meas_run_sts = {True: "Started", False: "Not Started"}
            self.application.measurement.meas_start_stop_timeout = timeout
            if not self.application.measurement.running:
                self.application.measurement.start()
                if not self.application.measurement.running:
                    self.__log.debug(f'⏳ waiting for measurement to start...')
                    self.application.measurement.wait_for_canoe_meas_to_start()
                self.__log.debug(f'👉 CANoe Measurement {meas_run_sts[self.get_measurement_running_status()]}.')
            else:
                self.__log.warning(f'⚠️ CANoe Measurement Already {meas_run_sts[self.application.measurement.running]}.')
            return self.application.measurement.running
        except Exception as e:
            self.__log.error(f'😡 Error starting measurement: {str(e)}')
            sys.exit(1)

    def stop_measurement(self, timeout=60) -> bool:
        """Stops the measurement.

        Args:
            timeout (int, optional): measurement start/stop event timeout in seconds. Defaults to 60.

        Returns:
            True if measurement stopped. else False.
        """
        return self.stop_ex_measurement(timeout)

    def stop_ex_measurement(self, timeout=60) -> bool:
        """StopEx repairs differences in the behavior of the Stop method on deferred stops concerning simulated and real mode in CANoe.

        Args:
            timeout (int, optional): measurement start/stop event timeout in seconds. Defaults to 60.

        Returns:
            True if measurement stopped. else False.
        """
        meas_run_sts = {True: "Not Stopped", False: "Stopped"}
        self.application.measurement.meas_start_stop_timeout = timeout
        if self.application.measurement.running:
            self.application.measurement.stop_ex()
            if self.application.measurement.running:
                self.__log.debug(f'⏳ waiting for measurement to stop...')
                self.application.measurement.wait_for_canoe_meas_to_stop()
            self.__log.debug(f'👉 CANoe Measurement {meas_run_sts[self.application.measurement.running]}.')
        else:
            self.__log.warning(f'⚠️ CANoe Measurement Already {meas_run_sts[self.application.measurement.running]}.')
        return not self.application.measurement.running

    def reset_measurement(self) -> bool:
        """reset(stop and start) the measurement.

        Returns:
            Measurement running status(True/False).
        """
        self.stop_ex_measurement()
        self.start_measurement()
        self.__log.debug(f'👉 measurement resetted on_demand.')
        return self.application.measurement.running

    def get_measurement_running_status(self) -> bool:
        """Returns the running state of the measurement.

        Returns:
            True if The measurement is running.
            False if The measurement is not running.
        """
        return self.application.measurement.running

    def add_offline_source_log_file(self, absolute_log_file_path: str) -> bool:
        """this method adds offline source log file.

        Args:
            absolute_log_file_path (str): absolute path of offline source log file.

        Returns:
            bool: returns True if log file added or already available. False if log file not available.
        """
        try:
            if os.path.isfile(absolute_log_file_path):
                # offline_sources = self.application.configuration.com_obj.OfflineSetup.Source.Sources
                offline_sources = self.application.configuration.offline_setup.source.sources.paths
                file_already_added = any([file == absolute_log_file_path for file in offline_sources])
                if file_already_added:
                    self.__log.warning(f'⚠️ offline logging file ({absolute_log_file_path}) already added.')
                else:
                    offline_sources.Add(absolute_log_file_path)
                    self.__log.debug(f'👉 added offline logging file ({absolute_log_file_path})')
                return True
            else:
                self.__log.debug(f'invalid logging file ({absolute_log_file_path}). Failed to add.')
                return False
        except Exception as e:
            self.__log.error(f'😡 Error adding offline logging file: {str(e)}')
            return False

    def start_measurement_in_animation_mode(self, animation_delay=100) -> None:
        """Starts the measurement in Animation mode.

        Args:
            animation_delay (int): The animation delay during the measurement in Offline Mode.
        """
        try:
            self.application.measurement.animation_delay = animation_delay
            self.application.measurement.animate()
            self.__log.debug(f"👉 started measurement in Animation mode with animation delay {animation_delay}.")
        except Exception as e:
            self.__log.error(f'😡 Error starting measurement in Animation mode: {str(e)}')

    def break_measurement_in_offline_mode(self) -> None:
        """Interrupts the playback in Offline mode."""
        try:
            self.application.measurement.break_offline_mode()
            self.__log.debug('👉 measurement interrupted in Offline mode.')
        except Exception as e:
            self.__log.error(f'😡 Error interrupting measurement in Offline mode: {str(e)}')

    def reset_measurement_in_offline_mode(self) -> None:
        """Resets the measurement in Offline mode."""
        try:
            self.application.measurement.reset_offline_mode()
            self.__log.debug('👉 measurement resetted in Offline mode.')
        except Exception as e:
            self.__log.error(f'😡 Error resetting measurement in Offline mode: {str(e)}')

    def step_measurement_event_in_single_step(self) -> None:
        """Processes a measurement event in single step."""
        try:
            self.application.measurement.step()
            self.__log.debug('👉 measurement event processed in single step.')
        except Exception as e:
            self.__log.error(f'😡 Error processing measurement event in single step: {str(e)}')

    def get_measurement_index(self) -> int:
        """gets the measurement index for the next measurement.

        Returns:
            Measurement Index.
        """
        self.__log.debug(f'👉 measurement_index value = {self.application.measurement.measurement_index}')
        return self.application.measurement.measurement_index

    def set_measurement_index(self, index: int) -> int:
        """sets the measurement index for the next measurement.

        Args:
            index (int): index value to set.

        Returns:
            Measurement Index value.
        """
        self.application.measurement.measurement_index = index
        self.__log.debug(f'👉 measurement_index value set to {index}')
        return self.application.measurement.measurement_index

    def save_configuration(self) -> bool:
        """Saves the configuration.

        Returns:
            True if configuration saved. else False.
        """
        try:
            if not self.application.configuration.saved:
                return self.application.configuration.save()
            else:
                self.__log.debug('😇 Active CANoe configuration already saved.')
                return self.application.configuration.saved
        except Exception as e:
            self.__log.error(f'😡 Error saving configuration: {str(e)}')
            return self.application.configuration.saved

    def save_configuration_as(self, path: str, major: int, minor: int, create_dir=True) -> bool:
        """Saves the configuration as a different CANoe version.

        Args:
            path (str): The complete file name.
            major (int): The major version number of the target version.
            minor (int): The minor version number of the target version.
            create_dir (bool): create directory if not available. default value True.

        Returns:
            True if configuration saved. else False.
        """
        try:
            config_path = '\\'.join(path.split('\\')[:-1])
            if not os.path.exists(config_path) and create_dir:
                os.makedirs(config_path, exist_ok=True)
            if os.path.exists(config_path):
                return self.application.configuration.save_as(path, major, minor, False)
            else:
                self.__log.error(f'😡 file path {config_path} not found.')
                return self.application.configuration.saved
        except Exception as e:
            self.__log.error(f'😡 Error saving configuration as: {str(e)}')
            return self.application.configuration.saved

    def get_can_bus_statistics(self, channel: int) -> dict:
        """Returns CAN Bus Statistics.

        Args:
            channel (int): The channel of the statistic that is to be returned.

        Returns:
            CAN bus statistics.
        """
        try:
            bus_types = {'CAN': 1, 'J1939': 2, 'TTP': 4, 'LIN': 5, 'MOST': 6, 'Kline': 14}
            can_bus_statistic_obj = self.application.configuration.online_setup.bus_statistics.bus_statistic(bus_types['CAN'], channel)
            statistics_info = {
                # The bus load
                'bus_load': can_bus_statistic_obj.bus_load,
                # The controller status
                'chip_state': can_bus_statistic_obj.chip_state,
                # The number of Error Frames per second
                'error': can_bus_statistic_obj.error,
                # The total number of Error Frames
                'error_total': can_bus_statistic_obj.error_total,
                # The number of messages with extended identifier per second
                'extended': can_bus_statistic_obj.extended,
                # The total number of messages with extended identifier
                'extended_total': can_bus_statistic_obj.extended_total,
                # The number of remote messages with extended identifier per second
                'extended_remote': can_bus_statistic_obj.extended_remote,
                # The total number of remote messages with extended identifier
                'extended_remote_total': can_bus_statistic_obj.extended_remote_total,
                # The number of overload frames per second
                'overload': can_bus_statistic_obj.overload,
                # The total number of overload frames
                'overload_total': can_bus_statistic_obj.overload_total,
                # The maximum bus load in 0.01 %
                'peak_load': can_bus_statistic_obj.peak_load,
                # Returns the current number of the Rx error counter
                'rx_error_count': can_bus_statistic_obj.rx_error_count,
                # The number of messages with standard identifier per second
                'standard': can_bus_statistic_obj.standard,
                # The total number of remote messages with standard identifier
                'standard_total': can_bus_statistic_obj.standard_total,
                # The number of remote messages with standard identifier per second
                'standard_remote': can_bus_statistic_obj.standard_remote,
                # The total number of remote messages with standard identifier
                'standard_remote_total': can_bus_statistic_obj.standard_remote_total,
                # The current number of the Tx error counter
                'tx_error_count': can_bus_statistic_obj.tx_error_count,
            }
            self.__log.debug(f'👉 CAN Bus Statistics 👉 {statistics_info}.')
            return statistics_info
        except Exception as e:
            self.__log.error(f'😡 Error getting CAN Bus Statistics: {str(e)}')
            return {}

    def get_canoe_version_info(self) -> dict:
        """The Version class represents the version of the CANoe application.

        Returns:
            "full_name" - The complete CANoe version.
            "name" - The CANoe version.
            "build" - The build number of the CANoe application.
            "major" - The major version number of the CANoe application.
            "minor" - The minor version number of the CANoe application.
            "patch" - The patch number of the CANoe application.
        """
        try:
            version_info = {'full_name': self.application.version.full_name,
                            'name': self.application.version.name,
                            'build': self.application.version.build,
                            'major': self.application.version.major,
                            'minor': self.application.version.minor,
                            'patch': self.application.version.patch}
            self.__log.debug('> CANoe Application.Version <'.center(100, '='))
            for k, v in version_info.items():
                self.__log.debug(f'{k:<10}: {v}')
            self.__log.debug(''.center(100, '='))
            return version_info
        except Exception as e:
            self.__log.error(f'😡 Error getting CANoe version info: {str(e)}')
            return {}

    def get_bus_databases_info(self, bus: str) -> dict:
        """returns bus database info(path, channel, full_name).

        Args:
            bus (str): bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.

        Returns:
            bus database info {'path': 'value', 'channel': 'value', 'full_name': 'value'}
        """
        try:
            dbcs_info = dict()
            databases_count = self.application.bus.databases.count
            for index in range(1, databases_count + 1):
                database_obj = self.application.bus.databases.database(index)
                dbcs_info[database_obj.name] = {'path': database_obj.path, 'channel': database_obj.channel, 'full_name': database_obj.full_name}
            self.__log.debug(f'👉 {bus} bus databases info -> {dbcs_info}.')
            return dbcs_info
        except Exception as e:
            self.__log.error(f'😡 Error getting {bus} bus databases info: {str(e)}')
            return {}

    def get_bus_nodes_info(self, bus: str) -> dict:
        """returns bus nodes info(path, full_name, active).

        Args:
            bus (str): bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.

        Returns:
            bus nodes info {'path': 'value', 'full_name': 'value', 'active': 'value'}
        """
        try:
            nodes_info = dict()
            nodes_count = self.application.bus.nodes.count
            for index in range(1, nodes_count + 1):
                node_obj = self.application.bus.nodes.node(index)
                nodes_info[node_obj.name] = {'path': node_obj.path, 'full_name': node_obj.full_name, 'active': node_obj.active}
            self.__log.debug(f'👉 {bus} bus nodes info -> {nodes_info}.')
            return nodes_info
        except Exception as e:
            self.__log.error(f'😡 Error getting {bus} bus nodes info: {str(e)}')
            return {}

    def get_signal_value(self, bus: str, channel: int, message: str, signal: str, raw_value=False) -> Union[int, float]:
        r"""get_signal_value Returns a Signal value.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.
            raw_value (bool): return raw value of the signal if true. Default(False) is physical value.

        Returns:
            signal value.
        """
        try:
            signal_obj = self.application.bus.get_signal(bus, channel, message, signal)
            signal_value = signal_obj.raw_value if raw_value else signal_obj.value
            self.__log.debug(f'👉 value of signal({bus}{channel}.{message}.{signal})={signal_value}.')
            return signal_value
        except Exception as e:
            self.__log.error(f'😡 Error getting signal value: {str(e)}')

    def set_signal_value(self, bus: str, channel: int, message: str, signal: str, value: Union[int, float], raw_value=False) -> None:
        r"""set_signal_value sets a value to Signal. Works only when messages are sent using CANoe IL.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.
            value (Union[float, int]): signal value.
            raw_value (bool): return raw value of the signal if true. Default(False) is physical value.
        """
        try:
            signal_obj = self.application.bus.get_signal(bus, channel, message, signal)
            if raw_value:
                signal_obj.raw_value = value
            else:
                signal_obj.value = value
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) value set to {value}.')
        except Exception as e:
            self.__log.error(f'😡 Error setting signal value: {str(e)}')

    def get_signal_full_name(self, bus: str, channel: int, message: str, signal: str) -> str:
        """Determines the fully qualified name of a signal.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.

        Returns:
            str: The fully qualified name of a signal. The following format will be used for signals: <DatabaseName>::<MessageName>::<SignalName>
        """
        try:
            signal_obj = self.application.bus.get_signal(bus, channel, message, signal)
            signal_fullname = signal_obj.full_name
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) full name = {signal_fullname}.')
            return signal_fullname
        except Exception as e:
            self.__log.error(f'😡 Error getting signal full name: {str(e)}')
            return ''

    def check_signal_online(self, bus: str, channel: int, message: str, signal: str) -> bool:
        r"""Checks whether the measurement is running and the signal has been received.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.

        Returns:
            TRUE if the measurement is running and the signal has been received. FALSE if not.
        """
        try:
            signal_obj = self.application.bus.get_signal(bus, channel, message, signal)
            sig_online_status = signal_obj.is_online
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) online status = {sig_online_status}.')
            return sig_online_status
        except Exception as e:
            self.__log.error(f'😡 Error checking signal online status: {str(e)}')
            return False

    def check_signal_state(self, bus: str, channel: int, message: str, signal: str) -> int:
        r"""Checks whether the measurement is running and the signal has been received.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.

        Returns:
            State of the signal.
                0- The default value of the signal is returned.
                1- The measurement is not running; the value set by the application is returned.
                2- The measurement is not running; the value of the last measurement is returned.
                3- The signal has been received in the current measurement; the current value is returned.
        """
        try:
            signal_obj = self.application.bus.get_signal(bus, channel, message, signal)
            sig_state = signal_obj.state
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) state = {sig_state}.')
            return sig_state
        except Exception as e:
            self.__log.error(f'😡 Error checking signal state: {str(e)}')

    def get_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, raw_value=False) -> Union[float, int]:
        r"""get_j1939_signal Returns a Signal object.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.
            source_addr (int): The source address of the ECU that sends the message.
            dest_addr (int): The destination address of the ECU that receives the message.
            raw_value (bool): return raw value of the signal if true. Default(False) is physical value.

        Returns:
            signal value.
        """
        try:
            signal_obj = self.application.bus.get_j1939_signal(bus, channel, message, signal, source_addr, dest_addr)
            signal_value = signal_obj.raw_value if raw_value else signal_obj.value
            self.__log.debug(f'👉 value of signal({bus}{channel}.{message}.{signal})={signal_value}.')
            return signal_value
        except Exception as e:
            self.__log.error(f'😡 Error getting signal value: {str(e)}')

    def set_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, value: Union[float, int], raw_value=False) -> None:
        r"""get_j1939_signal Returns a Signal object.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.
            source_addr (int): The source address of the ECU that sends the message.
            dest_addr (int): The destination address of the ECU that receives the message.
            value (Union[float, int]): signal value.
            raw_value (bool): return raw value of the signal if true. Default(False) is physical value.

        Returns:
            signal value.
        """
        try:
            signal_obj = self.application.bus.get_j1939_signal(bus, channel, message, signal, source_addr, dest_addr)
            if raw_value:
                signal_obj.raw_value = value
            else:
                signal_obj.value = value
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) value set to {value}.')
        except Exception as e:
            self.__log.error(f'😡 Error setting signal value: {str(e)}')

    def get_j1939_signal_full_name(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> str:
        """Determines the fully qualified name of a signal.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.
            source_addr (int): The source address of the ECU that sends the message.
            dest_addr (int): The destination address of the ECU that receives the message.

        Returns:
            str: The fully qualified name of a signal. The following format will be used for signals: <DatabaseName>::<MessageName>::<SignalName>
        """
        try:
            signal_obj = self.application.bus.get_j1939_signal(bus, channel, message, signal, source_addr, dest_addr)
            signal_fullname = signal_obj.full_name
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) full name = {signal_fullname}.')
            return signal_fullname
        except Exception as e:
            self.__log.error(f'😡 Error getting signal full name: {str(e)}')
            return ''

    def check_j1939_signal_online(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> bool:
        """Checks whether the measurement is running and the signal has been received.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.
            source_addr (int): The source address of the ECU that sends the message.
            dest_addr (int): The destination address of the ECU that receives the message.
        
        Returns:
            bool: TRUE: if the measurement is running and the signal has been received. FALSE: if not.
        """
        try:
            signal_obj = self.application.bus.get_j1939_signal(bus, channel, message, signal, source_addr, dest_addr)
            sig_online_status = signal_obj.is_online
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) online status = {sig_online_status}.')
            return sig_online_status
        except Exception as e:
            self.__log.error(f'😡 Error checking signal online status: {str(e)}')
            return False

    def check_j1939_signal_state(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> int:
        """Returns the state of the signal.

        Returns:
            int: State of the signal.
                possible values are:
                    0: The default value of the signal is returned.
                    1: The measurement is not running; the value set by the application is returned.
                    3: The signal has been received in the current measurement; the current value is returned.
        """
        try:
            signal_obj = self.application.bus.get_j1939_signal(bus, channel, message, signal, source_addr, dest_addr)
            sig_state = signal_obj.state
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) state = {sig_state}.')
            return sig_state
        except Exception as e:
            self.__log.error(f'😡 Error checking signal state: {str(e)}')

    def ui_activate_desktop(self, name: str) -> None:
        """Activates the desktop with the given name.

        Args:
            name (str): The name of the desktop to be activated.
        """
        try:
            self.application.ui.activate_desktop(name)
            self.__log.debug(f'👉 Activated the desktop with the given name({name}.')
        except Exception as e:
            self.__log.error(f'😡 Error activating the desktop: {str(e)}')

    def ui_open_baudrate_dialog(self) -> None:
        """opens the dialog for configuring the bus parameters. Make sure Measurement stopped when using this method."""
        try:
            self.application.ui.open_baudrate_dialog()
            self.__log.debug('👉 baudrate dialog opened. Configure the bus parameters.')
        except Exception as e:
            self.__log.error(f'😡 Error opening baudrate dialog: {str(e)}')

    def write_text_in_write_window(self, text: str) -> None:
        """Outputs a line of text in the Write Window.
        Args:
            text (str): The text.
        """
        try:
            self.application.ui.output_text_in_write_window(text)
            self.__log.debug(f'👉 text "{text}" written in the Write Window.')
        except Exception as e:
            self.__log.error(f'😡 Error writing text in the Write Window: {str(e)}')
    
    def read_text_from_write_window(self) -> str:
        """read the text contents from Write Window.

        Returns:
            The text content.
        """
        try:
            return self.application.ui.get_write_window_text
        except Exception as e:
            self.__log.error(f'😡 Error reading text from Write Window: {str(e)}')
            return ''
        
    def clear_write_window_content(self) -> None:
        """Clears the contents of the Write Window."""
        try:
            self.application.ui.clear_write_window()
            self.__log.debug('👉 Write Window content cleared.')
        except Exception as e:
            self.__log.error(f'😡 Error clearing Write Window content: {str(e)}')
    
    def copy_write_window_content(self) -> None:
        """Copies the contents of the Write Window to the clipboard."""
        try:
            self.application.ui.copy_write_window_content()
            self.__log.debug('👉 Write Window content copied to clipboard.')
        except Exception as e:
            self.__log.error(f'😡 Error copying Write Window content: {str(e)}')

    def enable_write_window_output_file(self, output_file: str, tab_index=None) -> None:
        """Enables logging of all outputs of the Write Window in the output file.

        Args:
            output_file (str): The complete path of the output file.
            tab_index (int, optional): The index of the page, for which logging of the output is to be activated. Defaults to None.
        """
        try:
            self.application.ui.enable_write_window_output_file(output_file, tab_index)
            self.__log.debug(f'👉 Enabled logging of outputs of the Write Window. output_file={output_file} and tab_index={tab_index}')
        except Exception as e:
            self.__log.error(f'😡 Error enabling Write Window output file: {str(e)}')
    
    def disable_write_window_output_file(self, tab_index=None) -> None:
        """Disables logging of all outputs of the Write Window.

        Args:
            tab_index (int, optional): The index of the page, for which logging of the output is to be activated. Defaults to None.
        """
        try:
            self.application.ui.disable_write_window_output_file(tab_index)
            self.__log.debug(f'👉 Disabled logging of outputs of the Write Window. tab_index={tab_index}')
        except Exception as e:
            self.__log.error(f'😡 Error disabling Write Window output file: {str(e)}')

    def define_system_variable(self, sys_var_name: str, value: Union[int, float, str]) -> object:
        """define_system_variable Create a system variable with an initial value
        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed"
            value (Union[int, float, str]): variable value.
        
        Returns:
            object: The new Variable object.
        """
        new_var_com_obj = None
        try:
            namespace_name = '::'.join(sys_var_name.split('::')[:-1])
            variable_name = sys_var_name.split('::')[-1]
            self.application.system.add_system_variable(namespace_name, variable_name, value)
            self.__log.debug(f'👉 system variable({sys_var_name}) created and value set to {value}.')
        except Exception as e:
            self.__log.error(f'😡 failed to create system variable({sys_var_name}). {e}')
        return new_var_com_obj
    
    def get_system_variable_value(self, sys_var_name: str) -> Union[int, float, str, tuple, None]:
        """get_system_variable_value Returns a system variable value.

        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed"

        Returns:
            System Variable value.
        """
        return_value = None
        try:            
            namespace = '::'.join(sys_var_name.split('::')[:-1])
            variable_name = sys_var_name.split('::')[-1]
            namespace_com_object = self.application.system.com_obj.Namespaces(namespace)
            variable_com_object = namespace_com_object.Variables(variable_name)
            return_value = variable_com_object.Value
            self.__log.debug(f'👉 system variable({sys_var_name}) value <- {return_value}.')
        except Exception as e:
            self.__log.debug(f'😡 failed to get system variable({sys_var_name}) value. {e}')
        return return_value
    
    def set_system_variable_value(self, sys_var_name: str, value: Union[int, float, str]) -> None:
        """set_system_variable_value sets a value to system variable.

        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed".
            value (Union[int, float, str]): variable value. supported CAPL system variable data types integer, double, string and data.
        """
        try:
            namespace = '::'.join(sys_var_name.split('::')[:-1])
            variable_name = sys_var_name.split('::')[-1]
            namespace_com_object = self.application.system.com_obj.Namespaces(namespace)
            variable_com_object = namespace_com_object.Variables(variable_name)
            if isinstance(variable_com_object.Value, int):
                variable_com_object.Value = int(value)
            elif isinstance(variable_com_object.Value, float):
                variable_com_object.Value = float(value)
            else:
                variable_com_object.Value = value
            self.__log.debug(f'👉 system variable({sys_var_name}) value set to -> {value}.')
        except Exception as e:
            self.__log.debug(f'😡 failed to set system variable({sys_var_name}) value. {e}')

    def set_system_variable_array_values(self, sys_var_name: str, value: tuple, index=0) -> None:
        """set_system_variable_array_values sets array of values to system variable.

        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed"
            value (tuple): variable values. supported integer array or double array. please always give only one type.
            index (int): value of index where values will start updating. Defaults to 0.
        """        
        try:
            namespace = '::'.join(sys_var_name.split('::')[:-1])
            variable_name = sys_var_name.split('::')[-1]
            namespace_com_object = self.application.system.com_obj.Namespaces(namespace)
            variable_com_object = namespace_com_object.Variables(variable_name)
            existing_variable_value = list(variable_com_object.Value)
            if (index + len(value)) <= len(existing_variable_value):
                final_value = existing_variable_value
                if isinstance(existing_variable_value[0], float):
                    final_value[index: index + len(value)] = (float(v) for v in value)
                else:
                    final_value[index: index + len(value)] = value
                variable_com_object.Value = tuple(final_value)
                wait(0.1)
                self.__log.debug(f'system variable({sys_var_name}) value set to -> {variable_com_object.Value}.')
            else:
                self.__log.debug(
                    f'failed to set system variable({sys_var_name}) value. check variable length and index value.')
        except Exception as e:
            self.__log.debug(f'failed to set system variable({sys_var_name}) value. {e}')

    def send_diag_request(self, diag_ecu_qualifier_name: str, request: str, request_in_bytes=True, return_sender_name=False) -> Union[str, dict]:
        r"""The send_diag_request method represents the query of a diagnostic tester (client) to an ECU (server) in CANoe.

        Args:
            diag_ecu_qualifier_name (str): Diagnostic Node ECU Qualifier Name configured in "Diagnostic/ISO TP Configuration".
            request (str): Diagnostic request in bytes or diagnostic request qualifier name.
            request_in_bytes (bool): True if Diagnostic request is bytes. False if you are using Qualifier name. Default is True.
            return_sender_name (bool): True if you user want response along with response sender name in dictionary. Default is False.

        Returns:
            diagnostic response stream. Ex- "50 01 00 00 00 00" or {'Door': "50 01 00 00 00 00"}
        """
        diag_response_data = ""
        diag_response_including_sender_name = {}
        try:
            if diag_ecu_qualifier_name in self.__diag_devices.keys():
                self.__log.debug(f'🚀 {diag_ecu_qualifier_name}: Diagnostic Request ➡️ {request}')
                if request_in_bytes:
                    diag_req = self.__diag_devices[diag_ecu_qualifier_name].create_request_from_stream(request)
                else:
                    diag_req = self.__diag_devices[diag_ecu_qualifier_name].create_request(request)
                diag_req.send()
                while diag_req.pending:
                    wait(0.1)
                diag_req_responses = diag_req.responses
                if len(diag_req_responses) == 0:
                    self.__log.debug("🙅 Diagnostic Response Not Received.")
                else:
                    for diag_res in diag_req_responses:
                        diag_response_data = diag_res.stream
                        diag_response_including_sender_name[diag_res.sender] = diag_response_data
                        if diag_res.positive:
                            self.__log.debug(f"🟢 {diag_res.sender}: Diagnostic Response ➕ve ⬅️ {diag_response_data}")
                        else:
                            self.__log.debug(f"🔴 {diag_res.Sender}: Diagnostic Response ➖ve ⬅️ {diag_response_data}")
            else:
                self.__log.warning(f'⚠️ Diagnostic ECU qualifier({diag_ecu_qualifier_name}) not available in loaded CANoe config.')
        except Exception as e:
            self.__log.error(f'😡 failed to send diagnostic request({request}). {e}')
        return diag_response_including_sender_name if return_sender_name else diag_response_data

    def control_tester_present(self, diag_ecu_qualifier_name: str, value: bool) -> None:
        """Starts/Stops sending autonomous/cyclical Tester Present requests to the ECU.

        Args:
            diag_ecu_qualifier_name (str): Diagnostic Node ECU Qualifier Name configured in "Diagnostic/ISO TP Configuration".
            value (bool): True - activate tester present. False - deactivate tester present.
        """
        try:
            if diag_ecu_qualifier_name in self.__diag_devices.keys():
                diag_device = self.__diag_devices[diag_ecu_qualifier_name]
                if diag_device.tester_present_status != value:
                    if value:
                        diag_device.start_tester_present()
                        self.__log.debug(f'⏱️🏃‍♂️ {diag_ecu_qualifier_name}: started tester present')
                    else:
                        diag_device.stop_tester_present()
                        self.__log.debug(f'⏱️🧍 {diag_ecu_qualifier_name}: stopped tester present')
                    wait(.1)
                else:
                    self.__log.warning(f'⚠️ {diag_ecu_qualifier_name}: tester present already set to {value}')
            else:
                self.__log.error(f'😇 diag ECU qualifier "{diag_ecu_qualifier_name}" not available in configuration.')
        except Exception as e:
            self.__log.error(f'😡 failed to control tester present. {e}')

    def set_replay_block_file(self, block_name: str, recording_file_path: str) -> None:
        r"""Method for setting CANoe replay block file.

        Args:
            block_name: CANoe replay block name
            recording_file_path: CANoe replay recording file including path.
        """
        try:
            replay_blocks = self.__replay_blocks
            if block_name in replay_blocks.keys():
                replay_block = replay_blocks[block_name]
                replay_block.path = recording_file_path
                self.__log.debug(f'👉 Replay block "{block_name}" updated with "{recording_file_path}" path.')
            else:
                self.__log.warning(f'⚠️ Replay block "{block_name}" not available.')
        except Exception as e:
            self.__log.error(f'😡 failed to set replay block file. {e}')

    def control_replay_block(self, block_name: str, start_stop: bool) -> None:
        r"""Method for setting CANoe replay block file.

        Args:
            block_name (str): CANoe replay block name
            start_stop (bool): True to start replay block. False to Stop.
        """
        try:
            replay_blocks = self.__replay_blocks
            if block_name in replay_blocks.keys():
                replay_block = replay_blocks[block_name]
                if start_stop:
                    replay_block.start()
                else:
                    replay_block.stop()
                self.__log.debug(f'👉 Replay block "{block_name}" {"Started" if start_stop else "Stopped"}.')
            else:
                self.__log.warning(f'⚠️ Replay block "{block_name}" not available.')
        except Exception as e:
            self.__log.error(f'😡 failed to control replay block. {e}')

    def compile_all_capl_nodes(self) -> dict:
        r"""compiles all CAPL, XML and .NET nodes.
        """
        try:
            capl_obj = self.application.capl
            capl_obj.compile()
            wait(1)
            compile_result = capl_obj.compile_result()
            self.__log.debug(f'👉 compiled all CAPL nodes successfully. result={compile_result["result"]}')
            return compile_result
        except Exception as e:
            self.__log.error(f'😡 failed to compile all CAPL nodes. {e}')
            return {}

    def call_capl_function(self, name: str, *arguments) -> bool:
        r"""Calls a CAPL function.
        Please note that the number of parameters must agree with that of the CAPL function.
        not possible to read return value of CAPL function at the moment. only execution status is returned.

        Args:
            name (str): The name of the CAPL function. Please make sure this name is already passed as argument during CANoe instance creation. see example for more info.
            arguments (tuple): Function parameters p1…p10 (optional).

        Returns:
            bool: CAPL function execution status. True-success, False-failed.
        """
        try:
            capl_obj = self.application.capl
            exec_sts = capl_obj.call_capl_function(self.application.measurement.user_capl_function_obj_dict[name], *arguments)
            self.__log.debug(f'👉 triggered capl function({name}). execution status = {exec_sts}.')
            return exec_sts
        except Exception as e:
            self.__log.error(f'😡 failed to call capl function({name}). {e}')
            return False

    def get_test_environments(self) -> dict:
        """returns dictionary of test environment names and class.
        """
        try:
            return self.__test_environments
        except Exception as e:
            self.__log.debug(f'😡 failed to get test environments. {e}')
            return {}
    
    def get_test_modules(self, test_env_name: str) -> dict:
        """returns dictionary of test module names and class.
        
        Args:
            test_env_name (str): test environment name. avoid duplicate test environment names in CANoe configuration.
        """
        try:
            test_environments = self.get_test_environments()
            if len(test_environments) > 0:
                if test_env_name in test_environments.keys():
                    return test_environments[test_env_name].get_all_test_modules()
                else:
                    self.__log.warning(f'⚠️ "{test_env_name}" not found in configuration.')
                    return {}
            else:
                self.__log.warning(f'⚠️ Zero test environments found in configuration. Not possible to fetch test modules.')
                return {}
        except Exception as e:
            self.__log.error(f'😡 failed to get test modules. {e}')
            return {}

    def execute_test_module(self, test_module_name: str) -> int:
        """use this method to execute test module.

        Args:
            test_module_name (str): test module name. avoid duplicate test module names in CANoe configuration.

        Returns:
            int: test module execution verdict. 0 ='VerdictNotAvailable', 1 = 'VerdictPassed', 2 = 'VerdictFailed',
        """
        try:
            test_verdict = {0: 'NotAvailable',
                            1: 'Passed',
                            2: 'Failed',
                            3: 'None (not available for test modules)',
                            4: 'Inconclusive (not available for test modules)',
                            5: 'ErrorInTestSystem (not available for test modules)', }
            execution_result = 0
            test_module_found = False
            test_env_name = ''
            for tm in self.__test_modules:
                if tm['name'] == test_module_name:
                    test_module_found = True
                    tm_obj = tm['object']
                    test_env_name = tm['environment']
                    self.__log.debug(f'👉 test module "{test_module_name}" found in "{test_env_name}"')
                    tm_obj.start()
                    tm_obj.wait_for_completion()
                    execution_result = tm_obj.verdict
                    break
                else:
                    continue
            if test_module_found and (execution_result == 1):
                self.__log.debug(f'👉 test module "{test_env_name}.{test_module_name}" executed and verdict = {test_verdict[execution_result]}.')
            elif test_module_found and (execution_result != 1):
                self.__log.debug(f'👉 test module "{test_env_name}.{test_module_name}" executed and verdict = {test_verdict[execution_result]}.')
            else:
                self.__log.warning(f'⚠️ test module "{test_env_name}.{test_module_name}" not found. not possible to execute.')
            return execution_result
        except Exception as e:
            self.__log.error(f'😡 failed to execute test module. {e}')
            return 0

    def stop_test_module(self, env_name: str, module_name: str):
        """stops execution of test module.
        
        Args:
            module_name (str): test module name. avoid duplicate test module names in CANoe configuration.
        """
        try:
            test_modules = self.get_test_modules(test_env_name=env_name)
            if test_modules:
                if module_name in test_modules.keys():
                    test_modules[module_name].stop()
                else:
                    self.__log.warning(f'⚠️ test module not found in "{env_name}" test environment.')
            else:
                self.__log.warning(f'⚠️ test modules not available in "{env_name}" test environment.')
        except Exception as e:
            self.__log.error(f'😡 failed to stop test module. {e}')

    def execute_all_test_modules_in_test_env(self, env_name: str):
        """executes all test modules available in test environment.
        
        Args:
            env_name (str): test environment name. avoid duplicate test environment names in CANoe configuration.
        """
        try:
            test_modules = self.get_test_modules(test_env_name=env_name)
            if test_modules:
                for tm_name in test_modules.keys():
                    self.execute_test_module(tm_name)
            else:
                self.__log.warning(f'⚠️ test modules not available in "{env_name}" test environment.')
        except Exception as e:
            self.__log.error(f'😡 failed to execute all test modules in "{env_name}" test environment. {e}')

    def stop_all_test_modules_in_test_env(self, env_name: str):
        """stops execution of all test modules available in test environment.
        
        Args:
            env_name (str): test environment name. avoid duplicate test environment names in CANoe configuration.
        """
        try:
            test_modules = self.get_test_modules(test_env_name=env_name)
            if test_modules:
                for tm_name in test_modules.keys():
                    self.stop_test_module(env_name, tm_name)
            else:
                self.__log.warning(f'⚠️ test modules not available in "{env_name}" test environment.')
        except Exception as e:
            self.__log.error(f'😡 failed to stop all test modules in "{env_name}" test environment. {e}')

    def execute_all_test_environments(self):
        """executes all test environments available in test setup.
        """
        try:
            test_environments = self.get_test_environments()
            if len(test_environments) > 0:
                for test_env_name in test_environments.keys():
                    self.__log.debug(f'👉 started executing test environment "{test_env_name}"...')
                    self.execute_all_test_modules_in_test_env(test_env_name)
                    self.__log.debug(f'👉 completed executing test environment "{test_env_name}"')
            else:
                self.__log.warning(f'⚠️ Zero test environments found in configuration.')
        except Exception as e:
            self.__log.error(f'😡 failed to execute all test environments. {e}')

    def stop_all_test_environments(self):
        """stops execution of all test environments available in test setup.
        """
        try:
            test_environments = self.get_test_environments()
            if len(test_environments) > 0:
                for test_env_name in test_environments.keys():
                    self.__log.debug(f'👉 stopping test environment "{test_env_name}" execution')
                    self.stop_all_test_modules_in_test_env(test_env_name)
                    self.__log.debug(f'👉 completed stopping test environment "{test_env_name}"')
            else:
                self.__log.warning(f'⚠️ Zero test environments found in configuration.')
        except Exception as e:
            self.__log.error(f'😡 failed to stop all test environments. {e}')

    def get_environment_variable_value(self, env_var_name: str) -> Union[int, float, str, tuple, None]:
        """returns a environment variable value.

        Args:
            env_var_name (str): The name of the environment variable. Ex- "float_var"

        Returns:
            Environment Variable value.
        """
        var_value = None
        try:
            variable = self.application.environment.get_variable(env_var_name)
            var_value = variable.value if variable.type != 3 else tuple(variable.value)
            self.__log.debug(f'👉 environment variable({env_var_name}) value ⬅️ {var_value}')
        except Exception as e:
            self.__log.error(f'😡 failed to get environment variable({env_var_name}) value. {e}')
        return var_value
    
    def set_environment_variable_value(self, env_var_name: str, value: Union[int, float, str, tuple]) -> None:
        r"""sets a value to environment variable.

        Args:
            env_var_name (str): The name of the environment variable. Ex- "speed".
            value (Union[int, float, str, tuple]): variable value. supported CAPL environment variable data types integer, double, string and data.
        """
        try:
            variable = self.application.environment.get_variable(env_var_name)
            if variable.type == 0:
                converted_value = int(value)
            elif variable.type == 1:
                converted_value = float(value)            
            elif variable.type == 2:
                converted_value = str(value)
            else:
                converted_value = tuple(value)
            variable.value = converted_value
            self.__log.debug(f'👉 environment variable({env_var_name}) value set to ➡️ {converted_value}')
        except Exception as e:
            self.__log.error(f'😡 failed to set system variable({env_var_name}) value. {e}')
