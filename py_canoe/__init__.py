# import external modules here
import os
import sys
import logging
import pythoncom
import win32com.client
from typing import Union
from datetime import datetime
from time import sleep as wait

# import internal modules here
from .py_canoe_logger import PyCanoeLogger


class CANoe:
    """
    Represents a CANoe instance.
    Args:
        py_canoe_log_dir (str): The path for the CANoe log file. Defaults to an empty string.
        user_capl_functions (tuple): A tuple of user-defined CAPL function names. Defaults to an empty tuple.
    """
    CANOE_APPLICATION_OPENED = False
    CANOE_APPLICATION_CLOSED = False
    CANOE_MEASUREMENT_STARTED = False
    CANOE_MEASUREMENT_STOPPED = False

    def __init__(self, py_canoe_log_dir='', user_capl_functions=tuple()):
        try:
            self.__log = PyCanoeLogger(py_canoe_log_dir).log
            self.application_events_enabled = True
            self.application_open_close_timeout = 60
            self.simulation_events_enabled = False
            self.measurement_events_enabled = True
            self.measurement_start_stop_timeout = 60   # default value set to 60 seconds (1 minute)
            self.configuration_events_enabled = False
            self.__user_capl_functions = user_capl_functions
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe object: {str(e)}')
            sys.exit(1)

    def __init_canoe_application(self):
        try:
            self.__log.debug('➖'*50)
            wait(0.5)
            pythoncom.CoInitialize()
            wait(0.5)
            self.application_com_obj = win32com.client.Dispatch('CANoe.Application')
            self.wait_for_canoe_app_to_open = lambda: DoMeasurementEventsUntil(lambda: CANoe.CANOE_APPLICATION_OPENED, lambda: self.application_open_close_timeout)
            self.wait_for_canoe_app_to_close = lambda: DoMeasurementEventsUntil(lambda: CANoe.CANOE_APPLICATION_CLOSED, lambda: self.application_open_close_timeout)
            if self.application_events_enabled:
                win32com.client.WithEvents(self.application_com_obj, CanoeApplicationEvents)
            wait(0.5)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe application: {str(e)}')
            sys.exit(1)

    def __init_canoe_application_bus(self):
        try:
            self.bus_com_obj = win32com.client.Dispatch(self.application_com_obj.Bus)
            self.bus_databases = win32com.client.Dispatch(self.bus_com_obj.Databases)
            self.bus_nodes = win32com.client.Dispatch(self.bus_com_obj.Nodes)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe bus: {str(e)}')
            sys.exit(1)

    def __init_canoe_application_capl(self):
        try:
            self.capl_obj = lambda: CanoeCapl(self.application_com_obj)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe CAPL: {str(e)}')
            sys.exit(1)

    def __init_canoe_application_configuration(self):
        try:
            self.configuration_com_obj = win32com.client.Dispatch(self.application_com_obj.Configuration)
            if self.configuration_events_enabled:
                win32com.client.WithEvents(self.configuration_com_obj, CanoeConfigurationEvents)
            self.configuration_offline_setup = win32com.client.Dispatch(self.configuration_com_obj.OfflineSetup)
            self.configuration_offline_setup_source = win32com.client.Dispatch(self.configuration_offline_setup.Source)
            self.configuration_offline_setup_source_sources = win32com.client.Dispatch(self.configuration_offline_setup_source.Sources)
            sources = self.configuration_offline_setup_source_sources
            sources_count = sources.Count + 1
            self.configuration_offline_setup_source_sources_paths = lambda: [sources.Item(index) for index in range(1, sources_count)]
            self.configuration_online_setup = win32com.client.Dispatch(self.configuration_com_obj.OnlineSetup)
            self.configuration_online_setup_bus_statistics = win32com.client.Dispatch(self.configuration_online_setup.BusStatistics)
            self.configuration_online_setup_bus_statistics_bus_statistic = lambda bus_type, channel: win32com.client.Dispatch(self.configuration_online_setup_bus_statistics.BusStatistic(bus_type, channel))
            self.configuration_general_setup = CanoeConfigurationGeneralSetup(self.configuration_com_obj)
            self.configuration_simulation_setup = lambda: CanoeConfigurationSimulationSetup(self.configuration_com_obj)
            self.__replay_blocks = self.configuration_simulation_setup().replay_collection.fetch_replay_blocks()
            self.configuration_test_setup = lambda: CanoeConfigurationTestSetup(self.configuration_com_obj)
            self.__test_setup_environments = self.configuration_test_setup().test_environments.fetch_all_test_environments()
            self.__test_modules = list()
            for te_name, te_inst in self.__test_setup_environments.items():
                for tm_name, tm_inst in te_inst.get_all_test_modules().items():
                    self.__test_modules.append({'name': tm_name, 'object': tm_inst, 'environment': te_name})
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe configuration: {str(e)}')

    def __init_canoe_application_environment(self):
        try:
            self.environment_obj_inst = CanoeEnvironment(self.application_com_obj)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe environment: {str(e)}')
            sys.exit(1)

    def __init_canoe_application_measurement(self):
        try:
            CanoeMeasurementEvents.application_com_obj = self.application_com_obj
            CanoeMeasurementEvents.user_capl_function_names = self.__user_capl_functions
            self.measurement_com_obj = win32com.client.Dispatch(self.application_com_obj.Measurement)
            self.wait_for_canoe_meas_to_start = lambda: DoMeasurementEventsUntil(lambda: CANoe.CANOE_MEASUREMENT_STARTED, lambda: self.measurement_start_stop_timeout)
            self.wait_for_canoe_meas_to_stop = lambda: DoMeasurementEventsUntil(lambda: CANoe.CANOE_MEASUREMENT_STOPPED, lambda: self.measurement_start_stop_timeout)
            if self.measurement_events_enabled:
                win32com.client.WithEvents(self.measurement_com_obj, CanoeMeasurementEvents)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe measurement: {str(e)}')
            sys.exit(1)

    def __init_canoe_application_networks(self):
        try:
            self.networks_com_obj = win32com.client.Dispatch(self.application_com_obj.Networks)
            self.networks_obj = lambda: CanoeNetworks(self.networks_com_obj)
            self.__diag_devices = self.networks_obj().fetch_all_diag_devices()
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe networks: {str(e)}')
            sys.exit(1)

    def __init_canoe_application_simulation(self):
        pass

    def __init_canoe_application_system(self):
        try:
            self.system_com_obj = win32com.client.Dispatch(self.application_com_obj.System)
            self.system_obj = lambda: CanoeSystem(self.system_com_obj)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe system: {str(e)}')
            sys.exit(1)

    def __init_canoe_application_ui(self):
        try:
            self.ui_com_obj = win32com.client.Dispatch(self.application_com_obj.UI)
            self.ui_write_window_com_obj = win32com.client.Dispatch(self.ui_com_obj.Write)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe UI: {str(e)}')
            sys.exit(1)

    def __init_canoe_application_version(self):
        try:
            self.version_com_obj = win32com.client.Dispatch(self.application_com_obj.Version)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe version: {str(e)}')
            sys.exit(1)

    def new(self, auto_save=False, prompt_user=False) -> None:
        try:
            self.__init_canoe_application()
            self.application_com_obj.New(auto_save, prompt_user)
            self.__log.debug(f'📢 New CANoe configuration successfully created 🎉')
        except Exception as e:
            self.__log.error(f'😡 Error creating new CANoe configuration: {str(e)}')
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
        self.__init_canoe_application()
        self.__init_canoe_application_measurement()
        self.__init_canoe_application_simulation()
        self.__init_canoe_application_version()
        try:
            self.application_com_obj.Visible = visible
            if self.measurement_com_obj.Running and not auto_stop:
                self.__log.error('😡 Measurement is running. Stop the measurement or set argument auto_stop=True')
                sys.exit(1)
            elif self.measurement_com_obj.Running and auto_stop:
                self.__log.warning('😇 Active Measurement is running. Stopping measurement before opening your configuration')
                self.stop_ex_measurement()
            if os.path.isfile(canoe_cfg):
                self.__log.debug('⏳ wait for application to open')
                self.application_com_obj.Open(canoe_cfg, auto_save, prompt_user)
                self.wait_for_canoe_app_to_open()
                self.__init_canoe_application_bus()
                self.__init_canoe_application_capl()
                self.__init_canoe_application_configuration()
                self.__init_canoe_application_environment()
                self.__init_canoe_application_networks()
                self.__init_canoe_application_system()
                self.__init_canoe_application_ui()
                self.__log.debug(f'📢 CANoe configuration successfully opened 🎉')
            else:
                self.__log.error(f'😡 CANoe configuration "{canoe_cfg}" not found')
                sys.exit(1)
        except Exception as e:
            self.__log.error(f'😡 Error opening CANoe configuration: {str(e)}')
            sys.exit(1)

    def quit(self):
        """Quits CANoe without saving changes in the configuration."""
        try:
            wait(0.5)
            self.__log.debug('⏳ wait for application to quit')
            self.application_com_obj.Quit()
            self.wait_for_canoe_app_to_close()
            wait(0.5)
            pythoncom.CoUninitialize()
            self.application_com_obj = None
            self.__log.debug('📢 CANoe Application Closed')
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
            meas_run_sts = {True: "Started 🏃‍♂️‍➡️", False: "Not Started 🧍‍♂️"}
            self.measurement_start_stop_timeout = timeout
            if self.measurement_com_obj.Running:
                self.__log.warning(f'⚠️ CANoe Measurement already running 🏃‍♂️‍➡️')
            else:
                self.measurement_com_obj.Start()
                if not self.measurement_com_obj.Running:
                    self.__log.debug(f'⏳ waiting for measurement to start')
                    self.wait_for_canoe_meas_to_start()
                    self.__log.debug(f'👉 CANoe Measurement {meas_run_sts[self.measurement_com_obj.Running]}')
            return self.measurement_com_obj.Running
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
        try:
            meas_run_sts = {True: "Not Stopped 🏃‍♂️‍➡️ ", False: "Stopped 🧍‍♂️"}
            self.measurement_start_stop_timeout = timeout
            if self.measurement_com_obj.Running:
                self.measurement_com_obj.Stop()
                if self.measurement_com_obj.Running:
                    self.__log.debug(f'⏳ waiting for measurement to stop 🧍‍♂️')
                    self.wait_for_canoe_meas_to_stop()
                    self.__log.debug(f'👉 CANoe Measurement {meas_run_sts[self.measurement_com_obj.Running]}')
            else:
                self.__log.warning(f'⚠️ CANoe Measurement already stopped 🧍‍♂️')
            return not self.measurement_com_obj.Running
        except Exception as e:
            self.__log.error(f'😡 Error stopping measurement: {str(e)}')
            sys.exit(1)

    def reset_measurement(self) -> bool:
        """reset(stop and start) the measurement.

        Returns:
            Measurement running status(True/False).
        """
        try:
            self.stop_measurement()
            self.start_measurement()
            self.__log.debug(f'👉 active measurement resetted 🔁')
            return self.measurement_com_obj.Running
        except Exception as e:
            self.__log.error(f'😡 Error resetting measurement: {str(e)}')
            sys.exit(1)

    def get_measurement_running_status(self) -> bool:
        """Returns the running state of the measurement.

        Returns:
            True if The measurement is running.
            False if The measurement is not running.
        """
        return self.measurement_com_obj.Running

    def add_offline_source_log_file(self, absolute_log_file_path: str) -> bool:
        """this method adds offline source log file.

        Args:
            absolute_log_file_path (str): absolute path of offline source log file.

        Returns:
            bool: returns True if log file added or already available. False if log file not available.
        """
        try:
            if os.path.isfile(absolute_log_file_path):
                offline_sources_paths = self.configuration_offline_setup_source_sources_paths()
                file_already_added = any([file == absolute_log_file_path for file in offline_sources_paths])
                if file_already_added:
                    self.__log.warning(f'⚠️ File "{absolute_log_file_path}" already added as offline source')
                else:
                    self.configuration_offline_setup_source_sources.Add(absolute_log_file_path)
                    self.__log.debug(f'📢 File "{absolute_log_file_path}" added as offline source')
                return True
            else:
                self.__log.error(f'😡 invalid logging file ({absolute_log_file_path})')
                return False
        except Exception as e:
            self.__log.error(f'😡 Error adding offline source log file: {str(e)}')
            return False

    def start_measurement_in_animation_mode(self, animation_delay=100) -> None:
        """Starts the measurement in Animation mode.

        Args:
            animation_delay (int): The animation delay during the measurement in Offline Mode.
        """
        try:
            self.measurement_com_obj.AnimationDelay = animation_delay
            self.measurement_com_obj.Animate()
            self.__log.debug(f'⏳ waiting for measurement to start 🏃‍♂️‍➡️')
            self.wait_for_canoe_meas_to_start()
            self.__log.debug(f"👉 started 🏃‍♂️‍➡️ measurement in Animation mode with animation delay ⏲️ {animation_delay}")
        except Exception as e:
            self.__log.error(f'😡 Error starting measurement in animation mode: {str(e)}')

    def break_measurement_in_offline_mode(self) -> None:
        """Interrupts the playback in Offline mode."""
        try:
            if self.measurement_com_obj.Running:
                self.measurement_com_obj.Break()
                self.__log.debug('👉 measurement interrupted 🫷 in Offline mode')
            else:
                self.__log.warning('⚠️ Measurement is not running')
        except Exception as e:
            self.__log.error(f'😡 Error interrupting measurement in Offline mode: {str(e)}')

    def reset_measurement_in_offline_mode(self) -> None:
        """Resets the measurement in Offline mode."""
        try:
            self.measurement_com_obj.Reset()
            self.__log.debug('👉 measurement resetted 🔁 in Offline mode')
        except Exception as e:
            self.__log.error(f'😡 Error resetting measurement in Offline mode: {str(e)}')

    def step_measurement_event_in_single_step(self) -> None:
        """Processes a measurement event in single step."""
        try:
            self.measurement_com_obj.Step()
            self.__log.debug('👉 Processed a measurement event in single step 👣')
        except Exception as e:
            self.__log.error(f'😡 Error stepping measurement in Single Step mode: {str(e)}')

    def get_measurement_index(self) -> int:
        """gets the measurement index for the next measurement.

        Returns:
            Measurement Index.
        """
        try:
            meas_index = self.measurement_com_obj.MeasurementIndex
            self.__log.debug(f'👉 measurement_index value 🟰 {meas_index}')
            return meas_index
        except Exception as e:
            self.__log.error(f'😡 Error getting measurement index: {str(e)}')
            return -1

    def set_measurement_index(self, index: int) -> int:
        """sets the measurement index for the next measurement.

        Args:
            index (int): index value to set.

        Returns:
            Measurement Index value.
        """
        try:
            self.measurement_com_obj.MeasurementIndex = index
            self.__log.debug(f'👉 measurement_index value set to ➡️ {index}')
            return index
        except Exception as e:
            self.__log.error(f'😡 Error setting measurement index: {str(e)}')
            return -1

    def save_configuration(self) -> bool:
        """Saves the configuration.

        Returns:
            True if configuration saved. else False.
        """
        try:
            if not self.configuration_com_obj.Saved:
                self.configuration_com_obj.Save()
                self.__log.debug('💾 configuration saved successfully')
            else:
                self.__log.debug('😇 configuration already saved')
            return self.configuration_com_obj.Saved
        except Exception as e:
            self.__log.error(f'😡 Error saving configuration: {str(e)}')
            return False

    def save_configuration_as(self, path: str, major: int, minor: int, prompt_user=False, create_dir=True) -> bool:
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
                self.configuration_com_obj.SaveAs(path, major, minor, prompt_user)
                if self.configuration_com_obj.Saved:
                    self.__log.debug(f'💾 configuration saved as {path} successfully')
                    return True
                else:
                    self.__log.error(f'😡 Error saving configuration as {path}')
                    return False
            else:
                self.__log.error(f'😡 file path {config_path} not found')
                return False
        except Exception as e:
            self.__log.error(f'😡 Error saving configuration as: {str(e)}')
            return False

    def get_can_bus_statistics(self, channel: int) -> dict:
        """Returns CAN Bus Statistics.

        Args:
            channel (int): The channel of the statistic that is to be returned.

        Returns:
            CAN bus statistics.
        """
        try:
            bus_types = {'CAN': 1, 'J1939': 2, 'TTP': 4, 'LIN': 5, 'MOST': 6, 'Kline': 14}
            can_bus_statistic_obj = self.configuration_online_setup_bus_statistics_bus_statistic(bus_types['CAN'], channel)
            statistics_info = {
                'bus_load': can_bus_statistic_obj.BusLoad,
                'chip_state': can_bus_statistic_obj.ChipState,
                'error': can_bus_statistic_obj.Error,
                'error_total': can_bus_statistic_obj.ErrorTotal,
                'extended': can_bus_statistic_obj.Extended,
                'extended_total': can_bus_statistic_obj.ExtendedTotal,
                'extended_remote': can_bus_statistic_obj.ExtendedRemote,
                'extended_remote_total': can_bus_statistic_obj.ExtendedRemoteTotal,
                'overload': can_bus_statistic_obj.Overload,
                'overload_total': can_bus_statistic_obj.OverloadTotal,
                'peak_load': can_bus_statistic_obj.PeakLoad,
                'rx_error_count': can_bus_statistic_obj.RxErrorCount,
                'standard': can_bus_statistic_obj.Standard,
                'standard_total': can_bus_statistic_obj.StandardTotal,
                'standard_remote': can_bus_statistic_obj.StandardRemote,
                'standard_remote_total': can_bus_statistic_obj.StandardRemoteTotal,
                'tx_error_count': can_bus_statistic_obj.TxErrorCount,
            }
            self.__log.debug(f'👉 CAN Bus Statistics ℹ️nfo 🟰 {statistics_info}')
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
            version_info = {'full_name': self.version_com_obj.FullName,
                            'name': self.version_com_obj.Name,
                            'build': self.version_com_obj.Build,
                            'major': self.version_com_obj.major,
                            'minor': self.version_com_obj.minor,
                            'patch': self.version_com_obj.Patch}
            self.__log.debug('> CANoe Application.Version ℹ️nfo<'.center(50, '➖'))
            for k, v in version_info.items():
                self.__log.debug(f'{k:<10}: {v}')
            self.__log.debug(''.center(50, '➖'))
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
            app_bus_databases_obj = win32com.client.Dispatch(self.application_com_obj.GetBus(bus).Databases)
            for item in app_bus_databases_obj:
                database_obj = win32com.client.Dispatch(item)
                dbcs_info[database_obj.Name] = {
                    'path': database_obj.Path,
                    'channel': database_obj.Channel,
                    'full_name': database_obj.FullName
                    }
            self.__log.debug(f'👉 {bus} bus databases ℹ️nfo 🟰 {dbcs_info}')
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
            app_bus_nodes_obj = win32com.client.Dispatch(self.application_com_obj.GetBus(bus).Nodes)
            for item in app_bus_nodes_obj:
                node_obj = win32com.client.Dispatch(item)
                nodes_info[node_obj.Name] = {
                    'path': node_obj.Path,
                    'full_name': node_obj.FullName,
                    'active': node_obj.Active
                    }
            self.__log.debug(f'👉 {bus} bus nodes ℹ️nfo 🟰 {nodes_info}')
            return nodes_info
        except Exception as e:
            self.__log.error(f'😡 Error getting {bus} bus nodes info: {str(e)}')
            return {}

    def get_signal_value(self, bus: str, channel: int, message: str, signal: str, raw_value=False) -> Union[int, float, None]:
        """get_signal_value Returns a Signal value.

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
            signal_obj = self.application_com_obj.GetBus(bus).GetSignal(channel, message, signal)
            signal_value = signal_obj.RawValue if raw_value else signal_obj.Value
            self.__log.debug(f'👉 value of signal({bus}{channel}.{message}.{signal}) 🟰 {signal_value}')
            return signal_value
        except Exception as e:
            self.__log.error(f'😡 Error getting signal value: {str(e)}')
            return None

    def set_signal_value(self, bus: str, channel: int, message: str, signal: str, value: Union[int, float], raw_value=False) -> None:
        """set_signal_value sets a value to Signal. Works only when messages are sent using CANoe IL.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.
            value (Union[float, int]): signal value.
            raw_value (bool): return raw value of the signal if true. Default(False) is physical value.
        """
        try:
            signal_obj = self.application_com_obj.GetBus(bus).GetSignal(channel, message, signal)
            if raw_value:
                signal_obj.RawValue = value
            else:
                signal_obj.Value = value
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) value set to {value}')
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
            signal_obj = self.application_com_obj.GetBus(bus).GetSignal(channel, message, signal)
            signal_fullname = signal_obj.FullName
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) full name 🟰 {signal_fullname}')
            return signal_fullname
        except Exception as e:
            self.__log.error(f'😡 Error getting signal full name: {str(e)}')
            return ''

    def check_signal_online(self, bus: str, channel: int, message: str, signal: str) -> bool:
        """Checks whether the measurement is running and the signal has been received.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.

        Returns:
            TRUE if the measurement is running and the signal has been received. FALSE if not.
        """
        try:
            signal_obj = self.application_com_obj.GetBus(bus).GetSignal(channel, message, signal)
            sig_online_status = signal_obj.IsOnline
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) online status 🟰 {sig_online_status}')
            return sig_online_status
        except Exception as e:
            self.__log.error(f'😡 Error checking signal online status: {str(e)}')
            return False

    def check_signal_state(self, bus: str, channel: int, message: str, signal: str) -> int:
        """Checks whether the measurement is running and the signal has been received.

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
            signal_obj = self.application_com_obj.GetBus(bus).GetSignal(channel, message, signal)
            sig_state = signal_obj.State
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) state 🟰 {sig_state}')
            return sig_state
        except Exception as e:
            self.__log.error(f'😡 Error checking signal state: {str(e)}')

    def get_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, raw_value=False) -> Union[float, int]:
        """get_j1939_signal Returns a Signal object.

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
            signal_obj = self.application_com_obj.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
            signal_value = signal_obj.RawValue if raw_value else signal_obj.Value
            self.__log.debug(f'👉 value of signal({bus}{channel}.{message}.{signal}) 🟰 {signal_value}')
            return signal_value
        except Exception as e:
            self.__log.error(f'😡 Error getting signal value: {str(e)}')

    def set_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, value: Union[float, int], raw_value=False) -> None:
        """get_j1939_signal Returns a Signal object.

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
            signal_obj = self.application_com_obj.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
            if raw_value:
                signal_obj.RawValue = value
            else:
                signal_obj.Value = value
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) value set to {value}')
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
            signal_obj = self.application_com_obj.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
            signal_fullname = signal_obj.FullName
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) full name 🟰 {signal_fullname}')
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
            signal_obj = self.application_com_obj.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
            sig_online_status = signal_obj.IsOnline
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) online status 🟰 {sig_online_status}')
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
            signal_obj = self.application_com_obj.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
            sig_state = signal_obj.State
            self.__log.debug(f'👉 signal({bus}{channel}.{message}.{signal}) state 🟰 {sig_state}')
            return sig_state
        except Exception as e:
            self.__log.error(f'😡 Error checking signal state: {str(e)}')

    def ui_activate_desktop(self, name: str) -> None:
        """Activates the desktop with the given name.

        Args:
            name (str): The name of the desktop to be activated.
        """
        try:
            self.ui_com_obj.ActivateDesktop(name)
            self.__log.debug(f'👉 Activated the desktop({name})')
        except Exception as e:
            self.__log.error(f'😡 Error activating the desktop: {str(e)}')

    def ui_open_baudrate_dialog(self) -> None:
        """opens the dialog for configuring the bus parameters. Make sure Measurement stopped when using this method."""
        try:
            self.ui_com_obj.OpenBaudrateDialog()
            self.__log.debug('👉 baudrate dialog opened. Configure the bus parameters')
        except Exception as e:
            self.__log.error(f'😡 Error opening baudrate dialog: {str(e)}')

    def write_text_in_write_window(self, text: str) -> None:
        """Outputs a line of text in the Write Window.
        Args:
            text (str): The text.
        """
        try:
            self.ui_write_window_com_obj.Output(text)
            self.__log.debug(f'✍️ text "{text}" written in the Write Window')
        except Exception as e:
            self.__log.error(f'😡 Error writing text in the Write Window: {str(e)}')

    def read_text_from_write_window(self) -> str:
        """read the text contents from Write Window.

        Returns:
            The text content.
        """
        try:
            text_content = self.ui_write_window_com_obj.Text
            self.__log.debug(f'📖 text read from Write Window: {text_content}')
            return text_content
        except Exception as e:
            self.__log.error(f'😡 Error reading text from Write Window: {str(e)}')
            return ''

    def clear_write_window_content(self) -> None:
        """Clears the contents of the Write Window."""
        try:
            self.ui_write_window_com_obj.Clear()
            self.__log.debug('🧹 Write Window content cleared')
        except Exception as e:
            self.__log.error(f'😡 Error clearing Write Window content: {str(e)}')

    def copy_write_window_content(self) -> None:
        """Copies the contents of the Write Window to the clipboard."""
        try:
            self.ui_write_window_com_obj.Copy()
            self.__log.debug('©️ Write Window content copied to clipboard')
        except Exception as e:
            self.__log.error(f'😡 Error copying Write Window content: {str(e)}')

    def enable_write_window_output_file(self, output_file: str, tab_index=None) -> None:
        """Enables logging of all outputs of the Write Window in the output file.

        Args:
            output_file (str): The complete path of the output file.
            tab_index (int, optional): The index of the page, for which logging of the output is to be activated. Defaults to None.
        """
        try:
            if tab_index:
                self.ui_write_window_com_obj.EnableOutputFile(output_file, tab_index)
                self.__log.debug(f'✔️ Enabled logging of outputs of the Write Window. output_file🟰{output_file} and tab_index🟰{tab_index}')
            else:
                self.ui_write_window_com_obj.EnableOutputFile(output_file)
                self.__log.debug(f'✔️ Enabled logging of outputs of the Write Window. output_file🟰{output_file}')
        except Exception as e:
            self.__log.error(f'😡 Error enabling Write Window output file: {str(e)}')

    def disable_write_window_output_file(self, tab_index=None) -> None:
        """Disables logging of all outputs of the Write Window.

        Args:
            tab_index (int, optional): The index of the page, for which logging of the output is to be activated. Defaults to None.
        """
        try:
            if tab_index:
                self.ui_write_window_com_obj.DisableOutputFile(tab_index)
                self.__log.debug(f'⏹️ Disabled logging of outputs of the Write Window. tab_index🟰{tab_index}')
            else:
                self.ui_write_window_com_obj.DisableOutputFile()
                self.__log.debug(f'⏹️ Disabled logging of outputs of the Write Window')
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
            system_obj = self.system_obj()
            system_obj.add_system_variable(namespace_name, variable_name, value)
            self.__log.debug(f'👉 system variable({sys_var_name}) created and value set to {value}')
        except Exception as e:
            self.__log.error(f'😡 failed to create system variable({sys_var_name}). {e}')
        return new_var_com_obj

    def get_system_variable_value(self, sys_var_name: str, return_symbolic_name=False) -> Union[int, float, str, tuple, dict, None]:
        """get_system_variable_value Returns a system variable value.

        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed"
            return_symbolic_name (bool): True if user want to return symbolic name. Default is False.

        Returns:
            System Variable value.
        """
        return_value = None
        try:
            namespace = '::'.join(sys_var_name.split('::')[:-1])
            variable_name = sys_var_name.split('::')[-1]
            namespace_com_object = self.system_com_obj.Namespaces(namespace)
            variable_com_object = win32com.client.Dispatch(namespace_com_object.Variables(variable_name))
            var_value = variable_com_object.Value
            if return_symbolic_name and (variable_com_object.Type == 0):
                var_value_name = variable_com_object.GetSymbolicValueName(var_value)
                return_value = var_value_name
            else:
                return_value = var_value
            self.__log.debug(f'👉 system variable({sys_var_name}) value 🟰 {return_value}')
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
            namespace_com_object = self.system_com_obj.Namespaces(namespace)
            variable_com_object = namespace_com_object.Variables(variable_name)
            if isinstance(variable_com_object.Value, int):
                variable_com_object.Value = int(value)
            elif isinstance(variable_com_object.Value, float):
                variable_com_object.Value = float(value)
            else:
                variable_com_object.Value = value
            self.__log.debug(f'👉 system variable({sys_var_name}) value set to {value}')
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
            namespace_com_object = self.system_com_obj.Namespaces(namespace)
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
                self.__log.debug(f'👉 system variable({sys_var_name}) value set to {variable_com_object.Value}')
            else:
                self.__log.warning(f'⚠️ failed to set system variable({sys_var_name}) value. check variable length and index value')
        except Exception as e:
            self.__log.error(f'😡 failed to set system variable({sys_var_name}) value. {e}')

    def send_diag_request(self, diag_ecu_qualifier_name: str, request: str, request_in_bytes=True, return_sender_name=False) -> Union[str, dict]:
        """The send_diag_request method represents the query of a diagnostic tester (client) to an ECU (server) in CANoe.

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
                self.__log.debug(f'💉 {diag_ecu_qualifier_name}: Diagnostic Request 🟰 {request}')
                if request_in_bytes:
                    diag_req = self.__diag_devices[diag_ecu_qualifier_name].create_request_from_stream(request)
                else:
                    diag_req = self.__diag_devices[diag_ecu_qualifier_name].create_request(request)
                diag_req.send()
                while diag_req.pending:
                    wait(0.1)
                diag_req_responses = diag_req.responses
                if len(diag_req_responses) == 0:
                    self.__log.warning("🙅 Diagnostic Response Not Received 🔴")
                else:
                    for diag_res in diag_req_responses:
                        diag_response_data = diag_res.stream
                        diag_response_including_sender_name[diag_res.sender] = diag_response_data
                        if diag_res.positive:
                            self.__log.debug(f"🟢 {diag_res.sender}: ➕ Diagnostic Response 👉 {diag_response_data}")
                        else:
                            self.__log.debug(f"🔴 {diag_res.Sender}: ➖ Diagnostic Response 👉 {diag_response_data}")
            else:
                self.__log.warning(f'⚠️ Diagnostic ECU qualifier({diag_ecu_qualifier_name}) not available in loaded CANoe config')
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
                        self.__log.debug(f'⏱️🏃‍♂️‍➡️‍ {diag_ecu_qualifier_name}: started tester present')
                    else:
                        diag_device.stop_tester_present()
                        self.__log.debug(f'⏱️🧍‍♂️ {diag_ecu_qualifier_name}: stopped tester present')
                    wait(.1)
                else:
                    self.__log.warning(f'⚠️ {diag_ecu_qualifier_name}: tester present already set to {value}')
            else:
                self.__log.error(f'😇 diag ECU qualifier "{diag_ecu_qualifier_name}" not available in configuration')
        except Exception as e:
            self.__log.error(f'😡 failed to control tester present. {e}')

    def set_replay_block_file(self, block_name: str, recording_file_path: str) -> None:
        """Method for setting CANoe replay block file.

        Args:
            block_name: CANoe replay block name
            recording_file_path: CANoe replay recording file including path.
        """
        try:
            replay_blocks = self.__replay_blocks
            if block_name in replay_blocks.keys():
                replay_block = replay_blocks[block_name]
                replay_block.path = recording_file_path
                self.__log.debug(f'👉 Replay block "{block_name}" updated with "{recording_file_path}" path')
            else:
                self.__log.warning(f'⚠️ Replay block "{block_name}" not available')
        except Exception as e:
            self.__log.error(f'😡 failed to set replay block file. {e}')

    def control_replay_block(self, block_name: str, start_stop: bool) -> None:
        """Method for controlling CANoe replay block.

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
                self.__log.debug(f'👉 Replay block "{block_name}" {"Started" if start_stop else "Stopped"}')
            else:
                self.__log.warning(f'⚠️ Replay block "{block_name}" not available')
        except Exception as e:
            self.__log.error(f'😡 failed to control replay block. {e}')

    def compile_all_capl_nodes(self) -> dict:
        """compiles all CAPL, XML and .NET nodes."""
        try:
            capl_obj = self.capl_obj()
            capl_obj.compile()
            wait(1)
            compile_result = capl_obj.compile_result()
            self.__log.debug(f'🧑‍💻 compiled all CAPL nodes successfully. result={compile_result["result"]}')
            return compile_result
        except Exception as e:
            self.__log.error(f'😡 failed to compile all CAPL nodes. {e}')
            return {}

    def call_capl_function(self, name: str, *arguments) -> bool:
        """Calls a CAPL function.
        Please note that the number of parameters must agree with that of the CAPL function.
        not possible to read return value of CAPL function at the moment. only execution status is returned.

        Args:
            name (str): The name of the CAPL function. Please make sure this name is already passed as argument during CANoe instance creation. see example for more info.
            arguments (tuple): Function parameters p1…p10 (optional).

        Returns:
            bool: CAPL function execution status. True-success, False-failed.
        """
        try:
            capl_obj = self.capl_obj()
            exec_sts = capl_obj.call_capl_function(CanoeMeasurementEvents.user_capl_function_obj_dict[name], *arguments)
            self.__log.debug(f'🛫 triggered capl function({name}). execution status 🟰 {exec_sts}')
            return exec_sts
        except Exception as e:
            self.__log.error(f'😡 failed to call capl function({name}). {e}')
            return False

    def get_test_environments(self) -> dict:
        """returns dictionary of test environment names and class."""
        try:
            return self.__test_setup_environments
        except Exception as e:
            self.__log.debug(f'😡 failed to get test environments. {e}')
            return {}

    def get_test_modules(self, env_name: str) -> dict:
        """returns dictionary of test environment test module names and its class object.

        Args:
            env_name (str): test environment name. avoid duplicate test environment names in CANoe configuration.
        """
        try:
            test_environments = self.get_test_environments()
            if len(test_environments) > 0:
                if env_name in test_environments.keys():
                    return test_environments[env_name].get_all_test_modules()
                else:
                    self.__log.warning(f'⚠️ "{env_name}" not found in configuration')
                    return {}
            else:
                self.__log.warning(f'⚠️ Zero test environments found in configuration. Not possible to fetch test modules')
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
                    self.__log.debug(f'🔎 test module "{test_module_name}" found in "{test_env_name}"')
                    tm_obj.start()
                    tm_obj.wait_for_completion()
                    execution_result = tm_obj.verdict
                    break
                else:
                    continue
            if test_module_found and (execution_result == 1):
                self.__log.debug(f'✔️ test module "{test_env_name}.{test_module_name}" executed and verdict 🟰 {test_verdict[execution_result]}')
            elif test_module_found and (execution_result != 1):
                self.__log.debug(f'😵‍💫 test module "{test_env_name}.{test_module_name}" executed and verdict 🟰 {test_verdict[execution_result]}')
            else:
                self.__log.warning(f'⚠️ test module "{test_module_name}" not found. not possible to execute')
            return execution_result
        except Exception as e:
            self.__log.error(f'😡 failed to execute test module. {e}')
            return 0

    def stop_test_module(self, test_module_name: str):
        """stops execution of test module.

        Args:
            test_module_name (str): test module name. avoid duplicate test module names in CANoe configuration.
        """
        try:
            for tm in self.__test_modules:
                if tm['name'] == test_module_name:
                    tm['object'].stop()
                    test_env_name = tm['environment']
                    self.__log.debug(f'👉 test module "{test_module_name}" in test environment "{test_env_name}" stopped 🧍‍♂️')
            else:
                self.__log.warning(f'⚠️ test module "{test_module_name}" not found. not possible to execute')
        except Exception as e:
            self.__log.error(f'😡 failed to stop test module. {e}')

    def execute_all_test_modules_in_test_env(self, env_name: str):
        """executes all test modules available in test environment.

        Args:
            env_name (str): test environment name. avoid duplicate test environment names in CANoe configuration.
        """
        try:
            test_modules = self.get_test_modules(env_name=env_name)
            if test_modules:
                for tm_name in test_modules.keys():
                    self.execute_test_module(tm_name)
            else:
                self.__log.warning(f'⚠️ test modules not available in "{env_name}" test environment')
        except Exception as e:
            self.__log.error(f'😡 failed to execute all test modules in "{env_name}" test environment. {e}')

    def stop_all_test_modules_in_test_env(self, env_name: str):
        """stops execution of all test modules available in test environment.

        Args:
            env_name (str): test environment name. avoid duplicate test environment names in CANoe configuration.
        """
        try:
            test_modules = self.get_test_modules(env_name=env_name)
            if test_modules:
                for tm_name in test_modules.keys():
                    self.stop_test_module(env_name, tm_name)
            else:
                self.__log.warning(f'⚠️ test modules not available in "{env_name}" test environment')
        except Exception as e:
            self.__log.error(f'😡 failed to stop all test modules in "{env_name}" test environment. {e}')

    def execute_all_test_environments(self):
        """executes all test environments available in test setup."""
        try:
            test_environments = self.get_test_environments()
            if len(test_environments) > 0:
                for test_env_name in test_environments.keys():
                    self.__log.debug(f'🏃‍♂️‍➡️ started executing test environment "{test_env_name}"')
                    self.execute_all_test_modules_in_test_env(test_env_name)
                    self.__log.debug(f'✔️ completed executing test environment "{test_env_name}"')
            else:
                self.__log.warning(f'⚠️ Zero test environments found in configuration')
        except Exception as e:
            self.__log.error(f'😡 failed to execute all test environments. {e}')

    def stop_all_test_environments(self):
        """stops execution of all test environments available in test setup."""
        try:
            test_environments = self.get_test_environments()
            if len(test_environments) > 0:
                for test_env_name in test_environments.keys():
                    self.__log.debug(f'⏹️ stopping test environment "{test_env_name}" execution')
                    self.stop_all_test_modules_in_test_env(test_env_name)
                    self.__log.debug(f'✔️ completed stopping test environment "{test_env_name}"')
            else:
                self.__log.warning(f'⚠️ Zero test environments found in configuration')
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
            variable = self.environment_obj_inst.get_variable(env_var_name)
            var_value = variable.value if variable.type != 3 else tuple(variable.value)
            self.__log.debug(f'👉 environment variable({env_var_name}) value 🟰 {var_value}')
        except Exception as e:
            self.__log.error(f'😡 failed to get environment variable({env_var_name}) value. {e}')
        return var_value

    def set_environment_variable_value(self, env_var_name: str, value: Union[int, float, str, tuple]) -> None:
        """sets a value to environment variable.

        Args:
            env_var_name (str): The name of the environment variable. Ex- "speed".
            value (Union[int, float, str, tuple]): variable value. supported CAPL environment variable data types integer, double, string and data.
        """
        try:
            variable = self.environment_obj_inst.get_variable(env_var_name)
            if variable.type == 0:
                converted_value = int(value)
            elif variable.type == 1:
                converted_value = float(value)
            elif variable.type == 2:
                converted_value = str(value)
            else:
                converted_value = tuple(value)
            variable.value = converted_value
            self.__log.debug(f'👉 environment variable({env_var_name}) value 🟰 {converted_value}')
        except Exception as e:
            self.__log.error(f'😡 failed to set system variable({env_var_name}) value. {e}')

    def add_database(self, database_file: str, database_network: str, database_channel: int) -> bool:
        try:
            if self.get_measurement_running_status():
                self.__log.warning('⚠️ measurement is running. not possible to add database')
                return False
            else:
                databases = self.configuration_general_setup.database_setup.databases.fetch_databases()
                if database_file in [database.full_name for database in databases.values()]:
                    self.__log.warning(f'⚠️ database "{database_file}" already added')
                    return False
                else:
                    self.configuration_general_setup.database_setup.databases.add_network(database_file, database_network)
                    wait(1)
                    databases = self.configuration_general_setup.database_setup.databases.fetch_databases()
                    for database in databases.values():
                        if database.full_name == database_file:
                            database.channel = database_channel
                            wait(1)
                    self.__log.debug(f'👉 database "{database_file}" added to network "{database_network}" and channel {database_channel}')
                    return True
        except Exception as e:
            self.__log.error(f'😡 failed to add database "{database_file}". {e}')
            return False

    def remove_database(self, database_file: str, database_channel: int) -> bool:
        try:
            if self.get_measurement_running_status():
                self.__log.warning('⚠️ measurement is running. not possible to remove database')
                return False
            else:
                databases = self.configuration_general_setup.database_setup.databases
                if database_file not in [database.full_name for database in databases.fetch_databases().values()]:
                    self.__log.warning(f'⚠️ database "{database_file}" not available to remove')
                    return False
                else:
                    for i in range(1, databases.count + 1):
                        database_com_obj = databases.com_obj.Item(i)
                        if database_com_obj.FullName == database_file and database_com_obj.Channel == database_channel:
                            self.configuration_general_setup.database_setup.databases.remove(i)
                            wait(1)
                            self.__log.debug(f'👉 database "{database_file}" removed from channel {database_channel}')
                            return True
        except Exception as e:
            self.__log.error(f'😡 failed to remove database "{database_file}". {e}')
            return False


def DoApplicationEvents() -> None:
    pythoncom.PumpWaitingMessages()
    wait(.1)

def DoApplicationEventsUntil(cond, timeout) -> None:
    base_time = datetime.now()
    while not cond():
        DoMeasurementEvents()
        now = datetime.now()
        difference = now - base_time
        seconds = difference.seconds
        if seconds > timeout():
            logging.getLogger('CANOE_LOG').debug(f'⌛ application event timeout({timeout()} s)')
            break

def DoMeasurementEvents() -> None:
    pythoncom.PumpWaitingMessages()
    wait(.1)

def DoMeasurementEventsUntil(cond, timeout) -> None:
    base_time = datetime.now()
    while not cond():
        DoMeasurementEvents()
        now = datetime.now()
        difference = now - base_time
        seconds = difference.seconds
        if seconds > timeout():
            logging.getLogger('CANOE_LOG').debug(f'⌛ measurement event timeout({timeout()} s)')
            break

def DoTestModuleEvents():
    pythoncom.PumpWaitingMessages()
    wait(.1)

def DoTestModuleEventsUntil(condition):
    while not condition():
        DoTestModuleEvents()

def DoEnvVarEvents():
    pythoncom.PumpWaitingMessages()
    wait(.1)

def DoEnvVarEventsUntil(condition):
    while not condition():
        DoEnvVarEvents()


class CanoeApplicationEvents:
    """Handler for CANoe Application events"""
    @staticmethod
    def OnOpen(fullname):
        CANoe.CANOE_APPLICATION_OPENED = True
        CANoe.CANOE_APPLICATION_CLOSED = False

    @staticmethod
    def OnQuit():
        CANoe.CANOE_APPLICATION_OPENED = False
        CANoe.CANOE_APPLICATION_CLOSED = True


class CanoeCapl:
    """The CAPL object allows to compile all nodes (CAPL, .NET, XML) in the configuration.
    Additionally, it represents the CAPL functions available in the CAPL programs.
    Please note that only user-defined CAPL functions can be accessed
    """
    def __init__(self, application_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(application_com_obj.CAPL)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CAPL object: {str(e)}')

    def compile(self) -> None:
        self.com_obj.Compile()

    def get_function(self, name: str) -> object:
        return self.com_obj.GetFunction(name)

    @staticmethod
    def parameter_count(capl_function_object: get_function) -> int:
        return capl_function_object.ParameterCount

    @staticmethod
    def parameter_types(capl_function_object: get_function) -> tuple:
        return capl_function_object.ParameterTypes

    def call_capl_function(self, capl_function_obj: get_function, *arguments) -> bool:
        return_value = False
        if len(arguments) == self.parameter_count(capl_function_obj):
            if len(arguments) > 0:
                capl_function_obj.Call(*arguments)
            else:
                capl_function_obj.Call()
            return_value = True
        else:
            self.__log.warning(fr'😇 function arguments not matching with CAPL user function args')
        return return_value

    def compile_result(self) -> dict:
        return_values = dict()
        compile_result_obj = self.com_obj.CompileResult
        return_values['error_message'] = compile_result_obj.ErrorMessage
        return_values['node_name'] = compile_result_obj.NodeName
        return_values['result'] = compile_result_obj.result
        return_values['source_file'] = compile_result_obj.SourceFile
        return return_values


class CanoeConfigurationEvents:
    """Handler for CANoe Configuration events"""
    @staticmethod
    def OnClose():
        logging.getLogger('CANOE_LOG').debug('👉 configuration OnClose event triggered')

    @staticmethod
    def OnSystemVariablesDefinitionChanged():
        logging.getLogger('CANOE_LOG').debug('👉 configuration OnSystemVariablesDefinitionChanged event triggered')


class CanoeConfigurationGeneralSetup:
    """The GeneralSetup object represents the general settings of a CANoe configuration."""
    def __init__(self, configuration_com_obj) -> None:
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(configuration_com_obj.GeneralSetup)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe general setup: {str(e)}')

    @property
    def database_setup(self):
        return CanoeConfigurationGeneralSetupDatabaseSetup(self.com_obj)


class CanoeConfigurationGeneralSetupDatabaseSetupEvents:
    @staticmethod
    def OnChange():
        logging.getLogger('CANOE_LOG').debug('👉 database setup OnChange event triggered')


class CanoeConfigurationGeneralSetupDatabaseSetup:
    """The DatabaseSetup object represents the assigned databases of the current configuration."""
    def __init__(self, general_setup_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(general_setup_com_obj.DatabaseSetup)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe database setup: {str(e)}')

    @property
    def databases(self):
        return CanoeConfigurationGeneralSetupDatabaseSetupDatabases(self.com_obj)


class CanoeConfigurationGeneralSetupDatabaseSetupDatabases:
    """The Databases object represents the assigned databases of CANoe."""
    def __init__(self, database_setup_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(database_setup_com_obj.Databases)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe databases: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def fetch_databases(self) -> dict:
        databases = dict()
        for index in range(1, self.count + 1):
            db_com_obj = self.com_obj.Item(index)
            db_inst = CanoeConfigurationGeneralSetupDatabaseSetupDatabasesDatabase(db_com_obj)
            databases[db_inst.name] = db_inst
        return databases

    def add(self, full_name: str) -> object:
        return self.com_obj.Add(full_name)

    def add_network(self, database_name: str, network_name: str) -> object:
        return self.com_obj.AddNetwork(database_name, network_name)

    def remove(self, index: int) -> None:
        self.com_obj.Remove(index)


class CanoeConfigurationGeneralSetupDatabaseSetupDatabasesDatabase:
    """The Database object represents the assigned database of the CANoe application."""
    def __init__(self, database_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(database_com_obj)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe database: {str(e)}')

    @property
    def channel(self) -> int:
        return self.com_obj.Channel

    @channel.setter
    def channel(self, channel: int) -> None:
        self.com_obj.Channel = channel

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @full_name.setter
    def full_name(self, full_name: str) -> None:
        self.com_obj.FullName = full_name

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def path(self) -> str:
        return self.com_obj.Path


class CanoeConfigurationSimulationSetup:
    """The SimulationSetup object represents the Simulation Setup of CANoe."""
    def __init__(self, configuration_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(configuration_com_obj.SimulationSetup)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe simulation setup: {str(e)}')

    @property
    def replay_collection(self):
        return CanoeConfigurationSimulationSetupReplayCollection(self.com_obj)

    @property
    def buses(self):
        return CanoeConfigurationSimulationSetupBuses(self.com_obj)

    @property
    def nodes(self):
        return CanoeConfigurationSimulationSetupNodes(self.com_obj)


class CanoeConfigurationSimulationSetupReplayCollection:
    """The ReplayCollection object represents the Replay Blocks of the CANoe application."""
    def __init__(self, sim_setup_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(sim_setup_com_obj.ReplayCollection)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe replay collection: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def add(self, name: str) -> object:
        return self.com_obj.Add(name)

    def remove(self, index: int) -> None:
        self.com_obj.Remove(index)

    def fetch_replay_blocks(self) -> dict:
        replay_blocks = dict()
        for index in range(1, self.count + 1):
            rb_com_obj = self.com_obj.Item(index)
            rb_inst = CanoeConfigurationSimulationSetupReplayCollectionReplayBlock(rb_com_obj)
            replay_blocks[rb_inst.name] = rb_inst
        return replay_blocks


class CanoeConfigurationSimulationSetupReplayCollectionReplayBlock:
    def __init__(self, replay_block_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(replay_block_com_obj)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe replay block: {str(e)}')

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def path(self) -> str:
        return self.com_obj.Path

    @path.setter
    def path(self, path: str):
        self.com_obj.Path = path

    def start(self):
        self.com_obj.Start()

    def stop(self):
        self.com_obj.Stop()


class CanoeConfigurationSimulationSetupBuses:
    """The Buses object represents the buses of the Simulation Setup of the CANoe application.
    The Buses object is only available in CANoe.
    """
    def __init__(self, sim_setup_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(sim_setup_com_obj.Buses)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe buses: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count


class CanoeConfigurationSimulationSetupNodes:
    """The Nodes object represents the CAPL node of the Simulation Setup of the CANoe application.
    The Nodes object is only available in CANoe.
    """
    def __init__(self, sim_setup_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(sim_setup_com_obj.Nodes)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe nodes: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count


class CanoeConfigurationTestSetup:
    """The TestSetup object represents CANoe's test setup."""
    def __init__(self, conf_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(conf_com_obj.TestSetup)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe test setup: {str(e)}')

    def save_all(self, prompt_user=False) -> None:
        self.com_obj.SaveAll(prompt_user)

    @property
    def test_environments(self):
        return CanoeConfigurationTestSetupTestEnvironments(self.com_obj)


class CanoeConfigurationTestSetupTestEnvironments:
    def __init__(self, test_setup_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(test_setup_com_obj.TestEnvironments)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe test environments: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def add(self, name: str) -> object:
        return self.com_obj.Add(name)

    def remove(self, index: int, prompt_user=False) -> None:
        self.com_obj.Remove(index, prompt_user)

    def fetch_all_test_environments(self) -> dict:
        test_environments = dict()
        for index in range(1, self.count + 1):
            te_com_obj = win32com.client.Dispatch(self.com_obj.Item(index))
            te_inst = CanoeConfigurationTestSetupTestEnvironmentsTestEnvironment(te_com_obj)
            test_environments[te_inst.name] = te_inst
        return test_environments


class CanoeConfigurationTestSetupTestEnvironmentsTestEnvironment:
    """The TestEnvironment object represents a test environment within CANoe's test setup."""
    def __init__(self, test_environment_com_obj):
        self.com_obj = test_environment_com_obj
        self.__test_modules = CanoeConfigurationTestSetupTestEnvironmentsTestEnvironmentTestModules(self.com_obj)

    @property
    def enabled(self) -> bool:
        return self.com_obj.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.com_obj.Enabled = value

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def path(self) -> str:
        return self.com_obj.Path

    def execute_all(self) -> None:
        self.com_obj.ExecuteAll()

    def save(self, name: str, prompt_user=False) -> None:
        self.com_obj.Save(name, prompt_user)

    def save_as(self, name: str, major: int, minor: int, prompt_user=False) -> None:
        self.com_obj.SaveAs(name, major, minor, prompt_user)

    def stop_sequence(self) -> None:
        self.com_obj.StopSequence()

    def get_all_test_modules(self):
        return self.__test_modules.fetch_test_modules()


class CanoeConfigurationTestSetupTestEnvironmentsTestEnvironmentTestModules:
    """The TestModules object represents the test modules in a test environment or in a test setup folder.
    This object should be preferred to the TestSetupItems object.
    """

    def __init__(self, test_env_com_obj) -> None:
        self.com_obj = test_env_com_obj.TestModules

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def add(self, full_name: str) -> object:
        return self.com_obj.Add(full_name)

    def remove(self, index: int, prompt_user=False) -> None:
        self.com_obj.Remove(index, prompt_user)

    def fetch_test_modules(self) -> dict:
        test_modules = dict()
        for index in range(1, self.count + 1):
            tm_com_obj = self.com_obj.Item(index)
            tm_inst = CanoeConfigurationTestSetupTestEnvironmentsTestEnvironmentTestModulesTestModule(tm_com_obj)
            test_modules[tm_inst.name] = tm_inst
        return test_modules


class CanoeConfigurationTestSetupTestEnvironmentsTestEnvironmentTestModulesTestModuleEvents:
    """test module events object."""
    def __init__(self):
        self.tm_html_report_path = ''
        self.tm_report_generated = False
        self.tm_running = False

    def OnStart(self):
        self.tm_html_report_path = ''
        self.tm_report_generated = False
        self.tm_running = True
        # logging.getLogger('CANOE_LOG').debug(f'👉test module OnStart event')

    @staticmethod
    def OnPause():
        logging.getLogger('CANOE_LOG').debug(f'👉test module OnPause event')

    def OnStop(self, reason):
        self.tm_running = False
        # logging.getLogger('CANOE_LOG').debug(f'👉test module OnStop event. reason -> {reason}')

    def OnReportGenerated(self, success, source_full_name, generated_full_name):
        self.tm_html_report_path = generated_full_name
        self.tm_report_generated = success
        self.tm_running = False
        logging.getLogger('CANOE_LOG').debug(f'👉test module OnReportGenerated event. {success} # {source_full_name} # {generated_full_name}')

    def OnVerdictFail(self):
        # logging.getLogger('CANOE_LOG').debug(f'👉test module OnVerdictFail event')
        pass


class CanoeConfigurationTestSetupTestEnvironmentsTestEnvironmentTestModulesTestModule:
    """The TestModule object represents a test module in CANoe's test setup."""

    def __init__(self, test_module_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.DispatchWithEvents(test_module_com_obj, CanoeConfigurationTestSetupTestEnvironmentsTestEnvironmentTestModulesTestModuleEvents)
            self.wait_for_tm_to_start = lambda: DoTestModuleEventsUntil(lambda: self.com_obj.tm_running)
            self.wait_for_tm_to_stop = lambda: DoTestModuleEventsUntil(lambda: not self.com_obj.tm_running)
            self.wait_for_tm_report_gen = lambda: DoTestModuleEventsUntil(lambda: self.com_obj.tm_report_generated)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe test module: {str(e)}')

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @property
    def path(self) -> str:
        return self.com_obj.Path

    @property
    def verdict(self) -> int:
        return self.com_obj.Verdict

    def start(self):
        self.com_obj.Start()
        self.wait_for_tm_to_start()
        logging.getLogger('CANOE_LOG').debug(f'👉 started executing test module. waiting for completion')

    def wait_for_completion(self):
        self.wait_for_tm_to_stop()
        wait(1)
        logging.getLogger('CANOE_LOG').debug(f'👉 completed executing test module. verdict = {self.verdict}')

    def pause(self) -> None:
        self.com_obj.Pause()

    def resume(self) -> None:
        self.com_obj.Resume()

    def stop(self) -> None:
        self.com_obj.Stop()
        logging.getLogger('CANOE_LOG').debug(f'👉stopping test module. waiting for completion')
        self.wait_for_tm_to_stop()
        logging.getLogger('CANOE_LOG').debug(f'👉completed stopping test module')

    def reload(self) -> None:
        self.com_obj.Reload()

    def set_execution_time(self, days: int, hours: int, minutes: int):
        self.com_obj.SetExecutionTime(days, hours, minutes)


class CanoeEnvironment:
    """The Environment class represents the environment variables.
    The Environment class is only available in CANoe
    """
    def __init__(self, application_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(application_com_obj.Environment)
        except Exception as e:
            self.__log.error(f'😡 Error initializing Environment object: {str(e)}')

    def create_group(self):
        return CanoeEnvironmentGroup(self.com_obj.CreateGroup())

    def create_info(self):
        return CanoeEnvironmentInfo(self.com_obj.CreateInfo())

    def get_variable(self, name: str):
        return CanoeEnvironmentVariable(self.com_obj.GetVariable(name))

    def get_variables(self, list_of_variable_names: tuple):
        return self.com_obj.GetVariables(list_of_variable_names)

    def set_variables(self, list_of_variables_with_name_value: tuple):
        self.com_obj.SetVariables(list_of_variables_with_name_value)


class CanoeEnvironmentGroup:
    """The EnvironmentGroup class represents a group of environment variables.
    With the help of environment variable groups you can set or query multiple environment variables simultaneously with just one call.
    """
    def __init__(self, env_group_com_obj):
        self.com_obj = env_group_com_obj

    @property
    def array(self):
        return CanoeEnvironmentArray(self.com_obj.Array)

    def add(self, variable):
        self.com_obj.Add(variable)

    def get_values(self):
        return self.com_obj.GetValues()

    def remove(self, variable):
        self.com_obj.Variable(variable)

    def set_values(self, values):
        self.com_obj.SetValues(values)


class CanoeEnvironmentArray:
    """The EnvironmentArray class represents an array of environment variables."""
    def __init__(self, env_array_com_obj):
        self.com_obj = env_array_com_obj

    @property
    def count(self) -> int:
        return self.com_obj.Count


class CanoeEnvironmentVariableEvents:
    def __init__(self):
        self.var_event_occurred = False

    def OnChange(self, value):
        self.var_event_occurred = True


class CanoeEnvironmentVariable:
    def __init__(self, env_var_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.DispatchWithEvents(env_var_com_obj, CanoeEnvironmentVariableEvents)
            self.wait_for_var_event = lambda: DoEnvVarEventsUntil(lambda: self.com_obj.var_event_occurred)
        except Exception as e:
            self.__log.error(f'😡 Error initializing EnvironmentVariable object: {str(e)}')

    @property
    def handle(self):
        return self.com_obj.Handle

    @handle.setter
    def handle(self, value):
        self.com_obj.Handle = value

    @property
    def notification_type(self):
        return self.com_obj.NotificationType

    @notification_type.setter
    def notification_type(self, value: int):
        self.com_obj.NotificationType = value

    @property
    def type(self):
        return self.com_obj.Type

    @property
    def value(self):
        return self.com_obj.Value

    @value.setter
    def value(self, value):
        self.com_obj.Value = value
        wait(.1)


class CanoeEnvironmentInfo:
    def __init__(self, env_info_com_obj):
        self.com_obj = env_info_com_obj

    @property
    def read(self):
        return self.com_obj.Read

    @property
    def write(self):
        return self.com_obj.Write

    def get_info(self):
        return self.com_obj.GetInfo()


class CanoeMeasurementEvents:
    application_com_obj = object
    user_capl_function_names = tuple()
    user_capl_function_obj_dict = dict()

    @staticmethod
    def OnInit():
        application_com_obj_loc = CanoeMeasurementEvents.application_com_obj
        for fun in CanoeMeasurementEvents.user_capl_function_names:
            CanoeMeasurementEvents.user_capl_function_obj_dict[fun] = application_com_obj_loc.CAPL.GetFunction(fun)
        CANoe.CANOE_MEASUREMENT_STARTED = False
        CANoe.CANOE_MEASUREMENT_STOPPED = False


    @staticmethod
    def OnStart():
        CANoe.CANOE_MEASUREMENT_STARTED = True
        CANoe.CANOE_MEASUREMENT_STOPPED = False


    @staticmethod
    def OnStop():
        CANoe.CANOE_MEASUREMENT_STARTED = False
        CANoe.CANOE_MEASUREMENT_STOPPED = True

    @staticmethod
    def OnExit():
        CANoe.CANOE_MEASUREMENT_STARTED = False
        CANoe.CANOE_MEASUREMENT_STOPPED = False


class CanoeNetworks:
    """The Networks class represents the networks of CANoe."""
    def __init__(self, networks_com_obj):
        try:
            self.log = logging.getLogger('CANOE_LOG')
            self.com_obj = networks_com_obj
        except Exception as e:
            self.log.error(f'😡 Error initializing Networks object: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def fetch_all_networks(self) -> dict:
        networks = dict()
        for index in range(1, self.count + 1):
            network_com_obj = win32com.client.Dispatch(self.com_obj.Item(index))
            network = CanoeNetworksNetwork(network_com_obj)
            networks[network_com_obj.Name] = network
        return networks

    def fetch_all_diag_devices(self) -> dict:
        diag_devices = dict()
        networks = self.fetch_all_networks()
        if len(networks) > 0:
            for _, n_value in networks.items():
                devices = n_value.devices
                n_devices = devices.get_all_devices()
                if len(n_devices) > 0:
                    for d_name, d_value in n_devices.items():
                        if d_value.diagnostic is not None:
                            diag_devices[d_name] = d_value.diagnostic
        return diag_devices


class CanoeNetworksNetwork:
    """The Network class represents one single network of CANoe."""
    def __init__(self, network_com_obj):
        self.com_obj = network_com_obj

    @property
    def bus_type(self) -> int:
        return self.com_obj.BusType

    @property
    def devices(self) -> object:
        return CanoeNetworksNetworkDevices(self.com_obj)

    @property
    def name(self) -> str:
        return self.com_obj.Name


class CanoeNetworksNetworkDevices:
    """The Devices class represents all devices of CANoe."""
    def __init__(self, network_com_obj):
        self.com_obj = network_com_obj.Devices

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def get_all_devices(self) -> dict:
        devices = dict()
        for index in range(1, self.count + 1):
            device_com_obj = self.com_obj.Item(index)
            device = CanoeNetworksNetworkDevicesDevice(device_com_obj)
            devices[device.name] = device
        return devices


class CanoeNetworksNetworkDevicesDevice:
    """The Device class represents one single device of CANoe."""
    def __init__(self, device_com_obj):
        self.com_obj = device_com_obj

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def diagnostic(self):
        try:
            diag_com_obj = self.com_obj.Diagnostic
            return CanoeNetworksNetworkDevicesDeviceDiagnostic(diag_com_obj)
        except pythoncom.com_error:
            return None


class CanoeNetworksNetworkDevicesDeviceDiagnostic:
    """The Diagnostic class represents the diagnostic properties of an ECU on the bus or the basic diagnostic functionality of a CANoe network.
    It is identified by the ECU qualifier that has been specified for the loaded diagnostic description (CDD/ODX).
    """
    def __init__(self, diagnostic_com_obj):
        self.com_obj = diagnostic_com_obj

    @property
    def tester_present_status(self) -> bool:
        return self.com_obj.TesterPresentStatus

    def create_request(self, primitive_path: str):
        return CanoeNetworksNetworkDevicesDeviceDiagnosticRequest(self.com_obj.CreateRequest(primitive_path))

    def create_request_from_stream(self, byte_stream: str):
        diag_req_in_bytes = bytearray()
        byte_stream = ''.join(byte_stream.split(' '))
        for i in range(0, len(byte_stream), 2):
            diag_req_in_bytes.append(int(byte_stream[i:i + 2], 16))
        return CanoeNetworksNetworkDevicesDeviceDiagnosticRequest(self.com_obj.CreateRequestFromStream(diag_req_in_bytes))

    def start_tester_present(self) -> None:
        self.com_obj.DiagStartTesterPresent()

    def stop_tester_present(self) -> None:
        self.com_obj.DiagStopTesterPresent()


class CanoeNetworksNetworkDevicesDeviceDiagnosticRequest:
    """The DiagnosticRequest class represents the query of a diagnostic tester (client) to an ECU (server) in CANoe.
    It can be replied by a DiagnosticResponse object.
    """
    def __init__(self, diag_req_com_obj):
        self.com_obj = diag_req_com_obj

    @property
    def pending(self) -> bool:
        return self.com_obj.Pending

    @property
    def responses(self) -> list:
        diag_responses_com_obj = self.com_obj.Responses
        diag_responses = [CanoeNetworksNetworkDevicesDeviceDiagnosticResponse(diag_responses_com_obj.item(i)) for i in range(1, diag_responses_com_obj.Count + 1)]
        return diag_responses

    @property
    def suppress_positive_response(self):
        return self.com_obj.SuppressPositiveResponse

    def send(self):
        self.com_obj.Send()

    def set_complex_parameter(self, qualifier, iteration, sub_parameter, value):
        self.com_obj.SetComplexParameter(qualifier, iteration, sub_parameter, value)

    def set_parameter(self, qualifier, value):
        self.com_obj.SetParameter(qualifier, value)


class CanoeNetworksNetworkDevicesDeviceDiagnosticResponse:
    """The DiagnosticResponse class represents the ECU's reply to a diagnostic request in CANoe.
    The received parameters can be read out and processed.
    """
    def __init__(self, diag_res_com_obj):
        self.com_obj = diag_res_com_obj

    @property
    def positive(self) -> bool:
        return self.com_obj.Positive

    @property
    def response_code(self) -> int:
        return self.com_obj.ResponseCode

    @property
    def stream(self) -> str:
        diag_response_data = " ".join(f"{d:02X}" for d in self.com_obj.Stream).upper()
        return diag_response_data

    @property
    def sender(self) -> str:
        return self.com_obj.Sender

    def get_complex_iteration_count(self, qualifier):
        return self.com_obj.GetComplexIterationCount(qualifier)

    def get_complex_parameter(self, qualifier, iteration, sub_parameter, mode):
        return self.com_obj.GetComplexParameter(qualifier, iteration, sub_parameter, mode)

    def get_parameter(self, qualifier, mode):
        return self.com_obj.GetParameter(qualifier, mode)

    def is_complex_parameter(self, qualifier):
        return self.com_obj.IsComplexParameter(qualifier)


class CanoeSystem:
    """The System object represents the system of the CANoe application.
    The System object offers access to the namespaces for data exchange with external applications.
    """
    def __init__(self, system_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = system_com_obj
            self.namespaces_com_obj = win32com.client.Dispatch(self.com_obj.Namespaces)
            self.variables_files_com_obj = win32com.client.Dispatch(self.com_obj.VariablesFiles)
            self.namespaces_dict = {}
            self.variables_files_dict = {}
            self.variables_dict = {}
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe System: {str(e)}')

    @property
    def namespaces_count(self) -> int:
        return self.namespaces_com_obj.Count

    def fetch_namespaces(self) -> dict:
        if self.namespaces_count > 0:
            for index in range(1, self.namespaces_count + 1):
                namespace_com_obj = win32com.client.Dispatch(self.namespaces_com_obj.Item(index))
                namespace_name = namespace_com_obj.Name
                self.namespaces_dict[namespace_name] = namespace_com_obj
                if 'Namespaces' in dir(namespace_com_obj):
                    self.fetch_namespace_namespaces(namespace_com_obj, namespace_name)
                if 'Variables' in dir(namespace_com_obj):
                    self.fetch_namespace_variables(namespace_com_obj)
        return self.namespaces_dict

    def add_namespace(self, name: str):
        self.fetch_namespaces()
        if name not in self.namespaces_dict.keys():
            namespace_com_obj = self.namespaces_com_obj.Add(name)
            self.namespaces_dict[name] = namespace_com_obj
            self.__log.debug(f'➕ Added the new namespace ({name})')
            return namespace_com_obj
        else:
            self.__log.warning(f'⚠️ The given namespace ({name}) already exists')
            return None

    def remove_namespace(self, name: str) -> None:
        self.fetch_namespaces()
        if name in self.namespaces_list:
            self.namespaces_com_obj.Remove(name)
            self.fetch_namespaces()
            self.__log.debug(f'➖ Removed the namespace ({name}) from the collection')
        else:
            self.__log.warning(f'⚠️ The given namespace ({name}) does not exist')

    @property
    def variables_files_count(self) -> int:
        return self.variables_files_com_obj.Count

    def fetch_variables_files(self):
        if self.variables_files_count > 0:
            for index in range(1, self.variables_files_count + 1):
                variable_file_com_obj = self.variables_files_com_obj.Item(index)
                self.variables_files_dict[variable_file_com_obj.Name] = {'full_name': variable_file_com_obj.FullName,
                                                                         'path': variable_file_com_obj.Path,
                                                                         'index': index}
        return self.variables_files_dict

    def add_variables_file(self, variables_file: str):
        self.fetch_variables_files()
        if os.path.isfile(variables_file):
            self.variables_files_com_obj.Add(variables_file)
            self.fetch_variables_files()
            self.__log.debug(f'➕ Added the Variables file ({variables_file}) to the collection')
        else:
            self.__log.warning(f'⚠️ The given file ({variables_file}) does not exist')

    def remove_variables_file(self, variables_file_name: str):
        self.fetch_variables_files()
        if variables_file_name in self.variables_files_dict:
            self.variables_files_com_obj.Remove(variables_file_name)
            self.fetch_variables_files()
            self.__log.debug(f'➖ Removed the Variables file ({variables_file_name}) from the collection')
        else:
            self.__log.warning(f'⚠️ The given file ({variables_file_name}) does not exist')

    def fetch_namespace_namespaces(self, parent_namespace_com_obj, parent_namespace_name):
        namespaces_count = parent_namespace_com_obj.Namespaces.Count
        if namespaces_count > 0:
            for index in range(1, namespaces_count + 1):
                namespace_com_obj = win32com.client.Dispatch(parent_namespace_com_obj.Namespaces.Item(index))
                namespace_name = f'{parent_namespace_name}::{namespace_com_obj.Name}'
                self.namespaces_dict[namespace_name] = namespace_com_obj
                if 'Namespaces' in dir(namespace_com_obj):
                    self.fetch_namespace_namespaces(namespace_com_obj, namespace_name)
                if 'Variables' in dir(namespace_com_obj):
                    self.fetch_namespace_variables(namespace_com_obj)

    def fetch_namespace_variables(self, parent_namespace_com_obj):
        variables_count = parent_namespace_com_obj.Variables.Count
        if variables_count > 0:
            for index in range(1, variables_count + 1):
                variable_obj = CanoeSystemVariable(parent_namespace_com_obj.Variables.Item(index))
                self.variables_dict[variable_obj.full_name] = variable_obj

    def add_system_variable(self, namespace, variable, value):
        self.fetch_namespaces()
        if f'{namespace}::{variable}' in self.variables_dict.keys():
            self.__log.warning(f'⚠️ The given variable ({variable}) already exists in the namespace ({namespace})')
            return None
        else:
            self.add_namespace(namespace)
            return self.namespaces_dict[namespace].Variables.Add(variable, value)

    def remove_system_variable(self, namespace, variable):
        self.fetch_namespaces()
        if f'{namespace}::{variable}' not in self.variables_dict.keys():
            self.__log.warning(f'⚠️ The given variable ({variable}) already removed in the namespace ({namespace})')
            return None
        else:
            self.namespaces_dict[namespace].Variables.Remove(variable)


class CanoeSystemVariable:
    def __init__(self, variable_com_obj):
        try:
            self.com_obj = win32com.client.Dispatch(variable_com_obj)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe Variable: {str(e)}')

    @property
    def analysis_only(self) -> bool:
        return self.com_obj.AnalysisOnly

    @analysis_only.setter
    def analysis_only(self, value: bool) -> None:
        self.com_obj.AnalysisOnly = value

    @property
    def bit_count(self) -> int:
        return self.com_obj.BitCount

    @property
    def comment(self) -> str:
        return self.com_obj.Comment

    @comment.setter
    def comment(self, text: str) -> None:
        self.com_obj.Comment = text

    @property
    def element_count(self) -> int:
        return self.com_obj.ElementCount

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @full_name.setter
    def full_name(self, full_name: str) -> None:
        self.com_obj.FullName = full_name

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def init_value(self) -> tuple[int, float, str]:
        return self.com_obj.InitValue

    @property
    def min_value(self) -> tuple[int, float, str]:
        return self.com_obj.MinValue

    @property
    def max_value(self) -> tuple[int, float, str]:
        return self.com_obj.MaxValue

    @property
    def is_array(self) -> bool:
        return self.com_obj.IsArray

    @property
    def is_signed(self) -> bool:
        return self.com_obj.IsSigned

    @property
    def read_only(self) -> bool:
        return self.com_obj.ReadOnly

    @property
    def type(self) -> int:
        return self.com_obj.Type

    @property
    def unit(self) -> str:
        return self.com_obj.Unit

    @property
    def value(self) -> tuple[int, float, str]:
        return self.com_obj.Value

    @value.setter
    def value(self, value: tuple[int, float, str]) -> None:
        self.com_obj.Value = value

    def get_member_phys_value(self, member_name: str):
        return self.com_obj.GetMemberPhysValue(member_name)

    def get_member_value(self, member_name: str):
        return self.com_obj.GetMemberValue(member_name)

    def get_symbolic_value_name(self, value: int):
        return self.com_obj.GetSymbolicValueName(value)

    def set_member_phys_value(self, member_name: str, value):
        return self.com_obj.setMemberPhysValue(member_name, value)

    def set_member_value(self, member_name: str, value):
        return self.com_obj.setMemberValue(member_name, value)

    def set_symbolic_value_name(self, value: int, name: str):
        self.com_obj.setSymbolicValueName(value, name)
