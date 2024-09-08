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
        self.__log = PyCanoeLogger(py_canoe_log_dir).log
        self.user_capl_function_names = user_capl_functions
        self.application = Application()

    def new(self, auto_save=False, prompt_user=False) -> None:
        """Creates a new configuration.

        Args:
            auto_save (bool, optional): A boolean value that indicates whether the active configuration should be saved if it has been changed. Defaults to False.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations. Defaults to False.
        """
        try:
            self.stop_ex_measurement()
            self.application.new(auto_save, prompt_user)
            self.__log.info('👉 created a new configuration.')
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
            self.__log.info(f'🔎 CANoe configuration "{canoe_cfg}" found.')
            self.application.open(canoe_cfg, auto_save, prompt_user)
            self.__log.info(f'📢 loaded CANoe configuration successfully 🎉')
        else:
            self.__log.error(f'😡 CANoe configuration "{canoe_cfg}" not found.')
            sys.exit(1)

    def quit(self):
        """Quits CANoe without saving changes in the configuration."""
        try:
            self.application.quit()
            self.__log.info('📢 CANoe Application Closed.')
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
                    self.__log.info(f'⏳ waiting for measurement to start...')
                    self.application.measurement.wait_for_canoe_meas_to_start()
                self.__log.info(f'👉 CANoe Measurement {meas_run_sts[self.get_measurement_running_status()]}.')
            else:
                self.__log.info(f'😇 CANoe Measurement Already {meas_run_sts[self.application.measurement.running]}.')
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
                self.__log.info(f'⏳ waiting for measurement to stop...')
                self.application.measurement.wait_for_canoe_meas_to_stop()
            self.__log.info(f'👉 CANoe Measurement {meas_run_sts[self.application.measurement.running]}.')
        else:
            self.__log.info(f'😇 CANoe Measurement Already {meas_run_sts[self.application.measurement.running]}.')
        return not self.application.measurement.running

    def reset_measurement(self) -> bool:
        """reset(stop and start) the measurement.

        Returns:
            Measurement running status(True/False).
        """
        self.stop_ex_measurement()
        self.start_measurement()
        self.__log.info(f'👉 measurement resetted on_demand.')
        return self.application.measurement.running

    def get_measurement_running_status(self) -> bool:
        """Returns the running state of the measurement.

        Returns:
            True if The measurement is running.
            False if The measurement is not running.
        """
        self.__log.info(f'👉 CANoe Measurement Running Status = {self.application.measurement.running}')
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
                    self.__log.info(f'😇 offline logging file ({absolute_log_file_path}) already added.')
                else:
                    offline_sources.Add(absolute_log_file_path)
                    self.__log.info(f'👉 added offline logging file ({absolute_log_file_path})')
                return True
            else:
                self.__log.info(f'invalid logging file ({absolute_log_file_path}). Failed to add.')
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
            self.__log.info('👉 started measurement in Animation mode.')
        except Exception as e:
            self.__log.error(f'😡 Error starting measurement in Animation mode: {str(e)}')

    def break_measurement_in_offline_mode(self) -> None:
        """Interrupts the playback in Offline mode."""
        try:
            self.application.measurement.break_offline_mode()
            self.__log.info('👉 measurement interrupted in Offline mode.')
        except Exception as e:
            self.__log.error(f'😡 Error interrupting measurement in Offline mode: {str(e)}')

    def reset_measurement_in_offline_mode(self) -> None:
        """Resets the measurement in Offline mode."""
        try:
            self.application.measurement.reset_offline_mode()
            self.__log.info('👉 measurement resetted in Offline mode.')
        except Exception as e:
            self.__log.error(f'😡 Error resetting measurement in Offline mode: {str(e)}')

    def step_measurement_event_in_single_step(self) -> None:
        """Processes a measurement event in single step."""
        try:
            self.application.measurement.step()
            self.__log.info('👉 measurement event processed in single step.')
        except Exception as e:
            self.__log.error(f'😡 Error processing measurement event in single step: {str(e)}')

    def get_measurement_index(self) -> int:
        """gets the measurement index for the next measurement.

        Returns:
            Measurement Index.
        """
        self.__log.info(f'👉 measurement_index value = {self.application.measurement.measurement_index}')
        return self.application.measurement.measurement_index

    def set_measurement_index(self, index: int) -> int:
        """sets the measurement index for the next measurement.

        Args:
            index (int): index value to set.

        Returns:
            Measurement Index value.
        """
        self.application.measurement.measurement_index = index
        self.__log.info(f'👉 measurement_index value set to {index}')
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
                self.__log.info('😇 Active CANoe configuration already saved.')
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
            self.__log.info(f'👉 CAN Bus Statistics 👉 {statistics_info}.')
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
            self.__log.info('> CANoe Application.Version <'.center(100, '='))
            for k, v in version_info.items():
                self.__log.info(f'{k:<10}: {v}')
            self.__log.info(''.center(100, '='))
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
            self.__log.info(f'👉 {bus} bus databases info -> {dbcs_info}.')
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
            self.__log.info(f'👉 {bus} bus nodes info -> {nodes_info}.')
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
            self.__log.info(f'👉 value of signal({bus}{channel}.{message}.{signal})={signal_value}.')
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
            self.__log.info(f'👉 signal({bus}{channel}.{message}.{signal}) value set to {value}.')
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
            self.__log.info(f'👉 signal({bus}{channel}.{message}.{signal}) full name = {signal_fullname}.')
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
            self.__log.info(f'👉 signal({bus}{channel}.{message}.{signal}) online status = {sig_online_status}.')
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
            self.__log.info(f'👉 signal({bus}{channel}.{message}.{signal}) state = {sig_state}.')
            return sig_state
        except Exception as e:
            self.__log.error(f'😡 Error checking signal state: {str(e)}')


# canoe_inst = CANoe()
# canoe_inst.open(r'C:\Users\Public\Documents\Vector\CANoe\Sample Configurations 15.5.23\CAN\CANSystemDemo\CANSystemDemo.cfg', auto_stop=True)
# canoe_inst.start_measurement()
# canoe_inst.stop_measurement()
# canoe_inst.reset_measurement()
# canoe_inst.get_measurement_running_status()
# canoe_inst.new()
# canoe_inst.quit()
# wait(5)
# canoe_inst.get_can_bus_statistics(1)
# canoe_inst.get_canoe_version_info()
# canoe_inst.get_bus_databases_info('CAN')
# canoe_inst.get_bus_nodes_info('CAN')
# canoe_inst.get_signal_value('CAN', 2, 'EngineData', 'EngSpeed')
# canoe_inst.set_signal_value('CAN', 2, 'EngineData', 'EngSpeed', 2000)
# wait(3)
# canoe_inst.get_signal_value('CAN', 2, 'EngineData', 'EngSpeed')
# canoe_inst.get_signal_full_name('CAN', 2, 'EngineData', 'EngSpeed')
# canoe_inst.check_signal_online('CAN', 2, 'EngineData', 'EngSpeed')
# canoe_inst.check_signal_state('CAN', 2, 'EngineData', 'EngSpeed')
# canoe_inst.stop_measurement()