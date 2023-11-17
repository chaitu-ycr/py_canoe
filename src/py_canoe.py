"""Python package for controlling Vector CANoe tool"""

# Import Python Libraries here
import os
import pythoncom
from typing import Union
from time import sleep as wait

# import CANoe utils here
from py_canoe_utils.py_canoe_logger import PyCanoeLogger
from py_canoe_utils.application import Application


class CANoe:
    r"""The CANoe class represents the CANoe application.
    The CANoe class is the foundation for the object hierarchy.
    You can reach all other methods from the CANoe class instance.

    Examples:
        >>> # Example to open CANoe configuration, start measurement, stop measurement and close configuration.
        >>> canoe_inst = CANoe(py_canoe_log_dir=r'D:\.py_canoe')
        >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
        >>> canoe_inst.start_measurement()
        >>> wait(10)
        >>> canoe_inst.stop_measurement()
        >>> canoe_inst.quit()
    """

    def __init__(self, py_canoe_log_dir='', user_capl_functions=tuple()):
        """
        Args:
            py_canoe_log_dir (str): directory to store py_canoe log. example 'D:\\.py_canoe'
            user_capl_functions (tuple): user defined CAPL functions to access. on measurement init these functions will be initialized.
        """
        pcl = PyCanoeLogger(py_canoe_log_dir)
        self.log = pcl.log
        self.application: Application
        self.__diag_devices = dict()
        self.__test_environments = dict()
        self.__test_modules = list()
        self.__replay_blocks = dict()
        self.user_capl_function_names = user_capl_functions

    def open(self, canoe_cfg: str, visible=True, auto_save=False, prompt_user=False) -> None:
        r"""Loads CANoe configuration.

        Args:
            canoe_cfg (str): The complete path for the CANoe configuration.
            visible (bool): True if you want to see CANoe UI. Defaults to True.
            auto_save (bool, optional): A boolean value that indicates whether the active configuration should be saved if it has been changed. Defaults to False.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations. Defaults to False.

        Examples:
            >>> # The following example opens a configuration
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
        """
        pythoncom.CoInitialize()
        self.application = Application(self.user_capl_function_names)
        self.application.visible = visible
        self.application.open(path=canoe_cfg, auto_save=auto_save, prompt_user=prompt_user)
        self.__diag_devices = self.application.networks.fetch_all_diag_devices()
        self.__test_environments = self.application.configuration.get_all_test_setup_environments()
        self.__test_modules = self.application.configuration.get_all_test_modules_in_test_environments()
        self.__replay_blocks = self.application.configuration.simulation_setup.replay_collection.fetch_replay_blocks()

    def new(self, auto_save=False, prompt_user=False) -> None:
        """Creates a new configuration.

        Args:
            auto_save (bool, optional): A boolean value that indicates whether the active configuration should be saved if it has been changed. Defaults to False.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations. Defaults to False.

        Examples:
            >>> # The following example creates a new configuration
            >>> canoe_inst = CANoe()
            >>> canoe_inst.new()
        """
        self.application.new(auto_save, prompt_user)

    def quit(self):
        r"""Quits CANoe without saving changes in the configuration.

        Examples:
            >>> # The following example quits CANoe
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.quit()
        """
        self.application.quit()

    def start_measurement(self) -> bool:
        r"""Starts the measurement.

        Returns:
            True if measurement started. else Flase.

        Examples:
            >>> # The following example starts the measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
        """
        return self.application.measurement.start()

    def stop_measurement(self) -> bool:
        r"""Stops the measurement.

        Returns:
            True if measurement stopped. else Flase.

        Examples:
            >>> # The following example stops the measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.stop_measurement()
        """
        return self.application.measurement.stop()

    def stop_ex_measurement(self) -> bool:
        r"""StopEx repairs differences in the behavior of the Stop method on deferred stops concerning simulated and real mode in CANoe.

        Returns:
            True if measurement stopped. else Flase.

        Examples:
            >>> # The following example full stops the measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.stop_ex_measurement()
        """
        return self.application.measurement.stop_ex()

    def reset_measurement(self) -> bool:
        r"""reset the measurement.

        Returns:
            Measurement running status(True/False).

        Examples:
            >>> # The following example resets the measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.reset_measurement()
        """
        if self.application.measurement.running:
            self.application.measurement.stop()
        self.application.measurement.start()
        self.log.info(f'Resetted measurement.')
        return self.application.measurement.running

    def get_measurement_running_status(self) -> bool:
        r"""Returns the running state of the measurement.

        Returns:
            True if The measurement is running.
            False if The measurement is not running.

        Examples:
            >>> # The following example returns measurement running status (True/False)
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.get_measurement_running_status()
        """
        self.log.info(f'CANoe Measurement Running Status = {self.application.measurement.running}')
        return self.application.measurement.running

    def add_offline_source_log_file(self, absolute_log_file_path: str) -> bool:
        r"""this method adds offline source log file.

        Args:
            absolute_log_file_path (str): absolute path of offline source log file.

        Returns:
            bool: returns True if log file added or already available. False if log file not available.
        
        Examples:
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.add_offline_source_log_file(fr'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\Logs\demo_log.blf')
        """
        if os.path.isfile(absolute_log_file_path):
            offline_sources = self.application.configuration.com_obj.OfflineSetup.Source.Sources
            file_already_added = any([file == absolute_log_file_path for file in offline_sources])
            if file_already_added:
                self.log.info(f'offline logging file ({absolute_log_file_path}) already added.')
            else:
                offline_sources.Add(absolute_log_file_path)
                self.log.info(f'added offline logging file ({absolute_log_file_path})')
            return True
        else:
            self.log.info(f'invalid logging file ({absolute_log_file_path}). Failed to add.')
            return False

    def start_measurement_in_animation_mode(self, animation_delay=100) -> None:
        r"""Starts the measurement in Animation mode.

        Args:
            animation_delay (int): The animation delay during the measurement in Offline Mode.

        Examples:
            >>> # The following example starts the measurement in Animation mode
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement_in_animation_mode()
        """
        self.application.measurement.animation_delay = animation_delay
        self.application.measurement.animate()

    def break_measurement_in_offline_mode(self) -> None:
        r"""Interrupts the playback in Offline mode.

        Examples:
            >>> # The following example interrupts the playback in Offline mode
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.break_measurement_in_offline_mode()
        """
        self.application.measurement.break_offline_mode()

    def reset_measurement_in_offline_mode(self) -> None:
        r"""Resets the measurement in Offline mode.

        Examples:
            >>> # The following example resets the measurement in Offline mode
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.reset_measurement_in_offline_mode()
        """
        self.application.measurement.reset_offline_mode()

    def step_measurement_event_in_single_step(self) -> None:
        r"""Processes a measurement event in single step.

        Examples:
            >>> # The following example processes a measurement event in single step
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.step_measurement_event_in_single_step()
        """
        self.application.measurement.step()

    def get_measurement_index(self) -> int:
        r"""gets the measurement index for the next measurement.

        Returns:
            Measurement Index.

        Examples:
            >>> # The following example gets the measurement index measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.stop_measurement()
            >>> canoe_inst.get_measurement_index()
        """
        self.log.info(f'measurement_index value = {self.application.measurement.measurement_index}')
        return self.application.measurement.measurement_index

    def set_measurement_index(self, index: int) -> int:
        r"""sets the measurement index for the next measurement.

        Args:
            index (int): index value to set.

        Returns:
            Measurement Index value.

        Examples:
            >>> # The following example sets the measurement index for the next measurement to 15
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.stop_measurement()
            >>> canoe_inst.set_measurement_index(15)
        """
        self.application.measurement.measurement_index = index
        return self.application.measurement.measurement_index

    def save_configuration(self) -> bool:
        r"""Saves the configuration.

        Returns:
            True if configuration saved. else False.

        Examples:
            >>> # The following example saves the configuration if necessary
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.save_configuration()
        """
        return self.application.configuration.save()

    def save_configuration_as(self, path: str, major: int, minor: int, create_dir=True) -> bool:
        r"""Saves the configuration as a different CANoe version.

        Args:
            path (str): The complete file name.
            major (int): The major version number of the target version.
            minor (int): The minor version number of the target version.
            create_dir (bool): create dirrectory if not available. default value True.

        Returns:
            True if configuration saved. else False.

        Examples:
            >>> # The following example saves the configuration as a CANoe 10.0 version
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.save_configuration_as(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo_v12.cfg', 10, 0)"""
        config_path = '\\'.join(path.split('\\')[:-1])
        if not os.path.exists(config_path) and create_dir:
            os.makedirs(config_path, exist_ok=True)
        if os.path.exists(config_path):
            self.application.configuration.save_as(path, major, minor, False)
            return self.application.configuration.saved
        else:
            self.log.info(f'tried creating {path}. but {config_path} directory not found.')
            return False

    def get_can_bus_statistics(self, channel: int) -> dict:
        r"""Returns CAN Bus Statistics.

        Args:
            channel (int): The channel of the statistic that is to be returned.

        Returns:
            CAN bus statistics.

        Examples:
            >>> # The following example prints CAN channel 1 statistics
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> print(canoe_inst.get_can_bus_statistics(channel=1))
        """
        conf_obj = self.application.configuration
        bus_types = {'CAN': 1, 'J1939': 2, 'TTP': 4, 'LIN': 5, 'MOST': 6, 'Kline': 14}
        bus_statistics_obj = conf_obj.com_obj.OnlineSetup.BusStatistics.BusStatistic(bus_types['CAN'], channel)
        statistics_info = {
            # The bus load
            'bus_load': bus_statistics_obj.BusLoad,
            # The controller status
            'chip_state': bus_statistics_obj.ChipState,
            # The number of Error Frames per second
            'error': bus_statistics_obj.Error,
            # The total number of Error Frames
            'error_total': bus_statistics_obj.ErrorTotal,
            # The number of messages with extended identifier per second
            'extended': bus_statistics_obj.Extended,
            # The total number of messages with extended identifier
            'extended_total': bus_statistics_obj.ExtendedTotal,
            # The number of remote messages with extended identifier per second
            'extended_remote': bus_statistics_obj.ExtendedRemote,
            # The total number of remote messages with extended identifier
            'extended_remote_total': bus_statistics_obj.ExtendedRemoteTotal,
            # The number of overload frames per second
            'overload': bus_statistics_obj.Overload,
            # The total number of overload frames
            'overload_total': bus_statistics_obj.OverloadTotal,
            # The maximum bus load in 0.01 %
            'peak_load': bus_statistics_obj.PeakLoad,
            # Returns the current number of the Rx error counter
            'rx_error_count': bus_statistics_obj.RxErrorCount,
            # The number of messages with standard identifier per second
            'standard': bus_statistics_obj.Standard,
            # The total number of remote messages with standard identifier
            'standard_total': bus_statistics_obj.StandardTotal,
            # The number of remote messages with standard identifier per second
            'standard_remote': bus_statistics_obj.StandardRemote,
            # The total number of remote messages with standard identifier
            'standard_remote_total': bus_statistics_obj.StandardRemoteTotal,
            # The current number of the Tx error counter
            'tx_error_count': bus_statistics_obj.TxErrorCount,
        }
        self.log.info(f'CAN Bus Statistics: {statistics_info}.')
        return statistics_info

    def get_canoe_version_info(self) -> dict:
        r"""The Version class represents the version of the CANoe application.

        Returns:
            "full_name" - The complete CANoe version.
            "name" - The CANoe version.
            "build" - The build number of the CANoe application.
            "major" - The major version number of the CANoe application.
            "minor" - The minor version number of the CANoe application.
            "patch" - The patch number of the CANoe application.

        Examples:
            >>> # The following example returns CANoe application version relevant information.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_version_info = canoe_inst.get_canoe_version_info()
            >>> print(canoe_version_info)
        """
        ver_obj = self.application.version
        version_info = {'full_name': ver_obj.full_name,
                        'name': ver_obj.name,
                        'build': ver_obj.build,
                        'major': ver_obj.major,
                        'minor': ver_obj.minor,
                        'patch': ver_obj.patch}
        self.log.info('> CANoe Application.Version <'.center(100, '='))
        for k, v in version_info.items():
            self.log.info(f'{k:<10}: {v}')
        self.log.info(''.center(100, '='))
        return version_info

    def get_bus_databases_info(self, bus: str):
        dbcs_info = dict()
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        db_objects = bus_obj.database_objects()
        for db_object in db_objects.values():
            dbcs_info[db_object.Name] = {'path': db_object.Path, 'channel': db_object.Channel,
                                         'full_name': db_object.FullName}
        self.log.info(f'{bus} bus databases info -> {dbcs_info}.')
        return dbcs_info

    def get_bus_nodes_info(self, bus: str):
        nodes_info = dict()
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        node_objects = bus_obj.node_objects()
        for n_object in node_objects.values():
            nodes_info[n_object.Name] = {'path': n_object.Path, 'full_name': n_object.FullName,
                                         'active': n_object.Active}
        self.log.info(f'{bus} bus nodes info -> {nodes_info}.')
        return nodes_info

    def get_signal_value(self, bus: str, channel: int, message: str, signal: str, raw_value=False) -> Union[float, int]:
        r"""get_signal_value Returns a Signal value.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.
            raw_value (bool): return raw value of the signal if true. Default(False) is physical value.

        Returns:
            signal vaue.

        Examples:
            >>> # The following example gets signal value
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> sig_val = canoe_inst.get_signal_value('CAN', 1, 'LightState', 'FlashLight')
            >>> print(sig_val)
        """
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        sig_obj = bus_obj.get_signal(channel, message, signal)
        signal_value = bus_obj.signal_get_raw_value(sig_obj) if raw_value else bus_obj.signal_get_value(sig_obj)
        self.log.info(f'value of signal({bus}{channel}.{message}.{signal})={signal_value}.')
        return signal_value

    def set_signal_value(self, bus: str, channel: int, message: str, signal: str, value: int, raw_value=False) -> None:
        r"""set_signal_value sets a value to Signal. Works only when messages are sent using CANoe IL.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.
            value (Union[float, int]): signal value.
            raw_value (bool): return raw value of the signal if true. Default(False) is physical value.

        Examples:
            >>> # The following example sets signal value to 1
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.set_signal_value('CAN', 1, 'LightState', 'FlashLight', 1)
        """
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        sig_obj = bus_obj.get_signal(channel, message, signal)
        if raw_value:
            bus_obj.signal_set_raw_value(sig_obj, value)
        else:
            bus_obj.signal_set_value(sig_obj, value)
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) value set to {value}.')

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
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        sig_obj = bus_obj.get_signal(channel, message, signal)
        signal_fullname = bus_obj.signal_full_name(sig_obj)
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) full name = {signal_fullname}.')
        return signal_fullname

    def check_signal_online(self, bus: str, channel: int, message: str, signal: str) -> bool:
        r"""Checks whether the measurement is running and the signal has been received.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.

        Returns:
            TRUE if the measurement is running and the signal has been received. FALSE if not.

        Examples:
            >>> # The following example checks signal is online.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.check_signal_online('CAN', 1, 'LightState', 'FlashLight')
        """
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        sig_obj = bus_obj.get_signal(channel, message, signal)
        sig_online_status = bus_obj.signal_is_online(sig_obj)
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) online status = {sig_online_status}.')
        return sig_online_status

    def check_signal_state(self, bus: str, channel: int, message: str, signal: str) -> int:
        r"""Checks whether the measurement is running and the signal has been received.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.

        Returns:
            State of the signal.
            0 The default value of the signal is returned.
            1 The measurement is not running; the value set by the application is returned.
            2 The measurement is not running; the value of the last measurement is returned.
            3 The signal has been received in the current measurement; the current value is returned.

        Examples:
            >>> # The following example checks signal state.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.check_signal_state('CAN', 1, 'LightState', 'FlashLight')
        """
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        sig_obj = bus_obj.get_signal(channel, message, signal)
        sig_state = bus_obj.signal_state(sig_obj)
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) state = {sig_state}.')
        return sig_state

    def get_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int,
                               dest_addr: int,
                               raw_value=False) -> Union[float, int]:
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
            signal vaue.

        Examples:
            >>> # The following example gets j1939 signal value
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> sig_val = canoe_inst.get_j1939_signal_value('CAN', 1, 'LightState', 'FlashLight', 0, 1)
            >>> print(sig_val)
        """
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        sig_obj = bus_obj.get_j1939_signal(channel, message, signal, source_addr, dest_addr)
        signal_value = bus_obj.signal_get_raw_value(sig_obj) if raw_value else bus_obj.signal_get_value(sig_obj)
        self.log.info(f'value of signal({bus}{channel}.{message}.{signal})={signal_value}.')
        return signal_value

    def set_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int,
                               dest_addr: int, value: Union[float, int],
                               raw_value=False) -> None:
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
            signal vaue.

        Examples:
            >>> # The following example gets j1939 signal value
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.set_j1939_signal_value('CAN', 1, 'LightState', 'FlashLight', 0, 1, 1)
        """
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        sig_obj = bus_obj.get_j1939_signal(channel, message, signal, source_addr, dest_addr)
        if raw_value:
            bus_obj.signal_set_raw_value(sig_obj, value)
        else:
            bus_obj.signal_set_value(sig_obj, value)
        self.log.info(f'signal value set to {value}.')
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) value set to {value}.')

    def get_j1939_signal_full_name(self, bus: str, channel: int, message: str, signal: str, source_addr: int,
                                   dest_addr: int) -> str:
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
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        sig_obj = bus_obj.get_j1939_signal(channel, message, signal, source_addr, dest_addr)
        signal_fullname = bus_obj.signal_full_name(sig_obj)
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) full name = {signal_fullname}.')
        return signal_fullname

    def check_j1939_signal_online(self, bus: str, channel: int, message: str, signal: str, source_addr: int,
                                  dest_addr: int) -> bool:
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
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        sig_obj = bus_obj.get_j1939_signal(channel, message, signal, source_addr, dest_addr)
        signal_online_status = bus_obj.signal_is_online(sig_obj)
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) online status = {signal_online_status}.')
        return signal_online_status

    def check_j1939_signal_state(self, bus: str, channel: int, message: str, signal: str, source_addr: int,
                                 dest_addr: int) -> int:
        """Returns the state of the signal.

        Returns:
            int: State of the signal.
                possible values are:
                    0: The default value of the signal is returned.
                    1: The measurement is not running; the value set by the application is returned.
                    3: The signal has been received in the current measurement; the current value is returned.
        """
        bus_obj = self.application.bus
        if bus_obj.bus_type != bus:
            bus_obj.reinit_bus(bus_type=bus)
        sig_obj = bus_obj.get_j1939_signal(channel, message, signal, source_addr, dest_addr)
        signal_state = bus_obj.signal_state(sig_obj)
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) state = {signal_state}.')
        return signal_state

    def ui_activate_desktop(self, name: str) -> None:
        r"""Activates the desktop with the given name.

        Args:
            name (str): The name of the desktop to be activated.

        Examples:
            >>> # The following example switches to the desktop with the name "Configuration"
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.ui_activate_desktop("Configuration")
        """
        self.application.ui.activate_desktop(name)

    def ui_open_baudrate_dialog(self) -> None:
        r"""opens the dialog for configuring the bus parameters. Make sure Measurement stopped when using this method.

        Examples:
            >>> # The following example opens the dialog for configuring the bus parameters
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.stop_measurement()
            >>> canoe_inst.ui_open_baudrate_dialog()
        """
        self.application.ui.open_baudrate_dialog()

    def write_text_in_write_window(self, text: str) -> None:
        r"""Outputs a line of text in the Write Window.
        Args:
            text (str): The text.

        Examples:
            >>> # The following example Outputs a line of text in the Write Window.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> canoe_inst.write_text_in_write_window("hello from python!")
            >>> wait(1)
            >>> print(canoe_inst.read_text_from_write_window())
        """
        self.application.ui.write.output(text)

    def read_text_from_write_window(self) -> str:
        r"""read the text contents from Write Window.

        Returns:
            The text content.

        Examples:
            >>> # The following example reads text from Write Window.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> canoe_inst.write_text_in_write_window("hello from python!")
            >>> wait(1)
            >>> print(canoe_inst.read_text_from_write_window())
        """
        return self.application.ui.write.text

    def clear_write_window_content(self) -> None:
        r"""Clears the contents of the Write Window.

        Examples:
            >>> # The following example clears content from Write Window.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> canoe_inst.write_text_in_write_window("hello from python!")
            >>> wait(1)
            >>> canoe_inst.clear_write_window_content()
        """
        self.application.ui.write.clear()

    def copy_write_window_content(self) -> None:
        r"""Copies the contents of the Write Window to the clipboard.

        Examples:
            >>> # The following example Copies the contents of the Write Window to the clipboard.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> canoe_inst.write_text_in_write_window("hello from python!")
            >>> wait(1)
            >>> canoe_inst.copy_write_window_content()
        """
        self.application.ui.write.copy()

    def enable_write_window_output_file(self, output_file: str, tab_index=None) -> None:
        r"""Enables logging of all outputs of the Write Window in the output file.

        Args:
            output_file (str): The complete path of the output file.
            tab_index (int, optional): The index of the page, for which logging of the output is to be activated. Defaults to None.

        Examples:
            >>> # The following example Enables logging of all outputs of the Write Window in the output file.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.enable_write_window_output_file(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\write_out.txt')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> canoe_inst.write_text_in_write_window("hello from python!")
            >>> wait(1)
            >>> canoe_inst.stop_measurement()
        """
        self.application.ui.write.enable_output_file(output_file, tab_index)

    def disable_write_window_output_file(self, tab_index=None) -> None:
        r"""Disables logging of all outputs of the Write Window.

        Args:
            tab_index (int, optional): The index of the page, for which logging of the output is to be activated. Defaults to None.

        Examples:
            >>> # The following example Disables logging of all outputs of the Write Window.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.disable_write_window_output_file()
        """
        self.application.ui.write.disable_output_file(tab_index)

    def define_system_variable(self, sys_var_name: str, value: Union[int, float, str]) -> object:
        r"""define_system_variable Create a system variable with an initial value
        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed"
            value (Union[int, float, str]): variable value. Default value 0.
        
        Returns:
            object: The new Variable object.
        
        Examples:
            >>> # The following example gets system variable value
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.define_system_variable('sys_demo::speed')
        """
        namespace_name = '::'.join(sys_var_name.split('::')[:-1])
        variable_name = sys_var_name.split('::')[-1]
        new_var_com_obj = None
        try:
            self.application.system.namespaces.add(namespace_name)
            namespaces = self.application.system.namespaces.fetch_namespaces()
            namespace = namespaces[namespace_name]
            new_var_com_obj = namespace.variables.add(variable_name, value)
            self.log.info(f'system variable({sys_var_name}) created and value set to {value}.')
        except Exception as e:
            self.log.info(f'failed to create system variable({sys_var_name}). {e}')
        return new_var_com_obj

    def get_system_variable_value(self, sys_var_name: str) -> Union[int, float, str, tuple, None]:
        r"""get_system_variable_value Returns a system variable value.

        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed"

        Returns:
            System Variable value.

        Examples:
            >>> # The following example gets system variable value
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> sys_var_val = canoe_inst.get_system_variable_value('sys_var_demo::speed')
            >>>print(sys_var_val)
        """
        namespace = '::'.join(sys_var_name.split('::')[:-1])
        variable_name = sys_var_name.split('::')[-1]
        return_value = None
        try:
            namespace_com_object = self.application.system.com_obj.Namespaces(namespace)
            variable_com_object = namespace_com_object.Variables(variable_name)
            return_value = variable_com_object.Value
            self.log.info(f'system variable({sys_var_name}) value <- {return_value}.')
        except Exception as e:
            self.log.info(f'failed to get system variable({sys_var_name}) value. {e}')
        return return_value

    def set_system_variable_value(self, sys_var_name: str, value: Union[int, float, str]) -> None:
        r"""set_system_variable_value sets a value to system variable.

        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed".
            value (Union[int, float, str]): variable value. supported CAPL system variable data types integer, double, string and data.

        Examples:
            >>> # The following example sets system variable value to 1
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.set_system_variable_value('sys_var_demo::speed', 1)
            >>> canoe_inst.set_system_variable_value('demo::string_var', 'hey hello this is string variable')
            >>> canoe_inst.set_system_variable_value('demo::data_var', 'hey hello this is data variable')
        """
        namespace = '::'.join(sys_var_name.split('::')[:-1])
        variable_name = sys_var_name.split('::')[-1]
        try:
            namespace_com_object = self.application.system.com_obj.Namespaces(namespace)
            variable_com_object = namespace_com_object.Variables(variable_name)
            if isinstance(variable_com_object.Value, int):
                variable_com_object.Value = int(value)
            elif isinstance(variable_com_object.Value, float):
                variable_com_object.Value = float(value)
            else:
                variable_com_object.Value = value
            self.log.info(f'system variable({sys_var_name}) value set to -> {value}.')
        except Exception as e:
            self.log.info(f'failed to set system variable({sys_var_name}) value. {e}')

    def set_system_variable_array_values(self, sys_var_name: str, value: tuple, index=0) -> None:
        r"""set_system_variable_array_values sets array of values to system variable.

        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed"
            value (tuple): variable values. supported integer array or double array. please always give only one type.
            index (int): value of index where values will start updating. Defaults to 0.

        Examples:
            >>> # The following example sets system variable value to 1
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.set_system_variable_array_values('demo::int_array_var', (00, 11, 22, 33, 44, 55, 66, 77, 88, 99))
            >>> canoe_inst.set_system_variable_array_values('demo::double_array_var', (00.0, 11.1, 22.2, 33.3, 44.4))
            >>> canoe_inst.set_system_variable_array_values('demo::double_array_var', (99.9, 100.0), 3)
        """
        namespace = '::'.join(sys_var_name.split('::')[:-1])
        variable_name = sys_var_name.split('::')[-1]
        try:
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
                self.log.info(f'system variable({sys_var_name}) value set to -> {variable_com_object.Value}.')
            else:
                self.log.info(
                    f'failed to set system variable({sys_var_name}) value. check variable length and index value.')
        except Exception as e:
            self.log.info(f'failed to set system variable({sys_var_name}) value. {e}')

    def send_diag_request(self, diag_ecu_qualifier_name: str, request: str, request_in_bytes=True,
                          return_sender_name=False) -> Union[str, dict]:
        r"""The send_diag_request method represents the query of a diagnostic tester (client) to an ECU (server) in CANoe.

        Args:
            diag_ecu_qualifier_name (str): Diagnostic Node ECU Qualifier Name configured in "Diagnostic/ISO TP Configuration".
            request (str): Diagnostic request in bytes or diagnostic request qualifier name.
            request_in_bytes (bool): True if Diagnostic request is bytes. False if you are using Qualifier name. Default is True.
            return_sender_name (bool): True if you user want response along with response sender name in dictionary. Default is False.

        Returns:
            diagnostic response stream. Ex- "50 01 00 00 00 00" or {'Door': "50 01 00 00 00 00"}

        Examples:
            >>> # Example 1 - The following example sends diagnostic request "10 01"
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> resp = canoe_inst.send_diag_request('Door', '10 01')
            >>> print(resp)
            >>> canoe_inst.stop_measurement()
            >>> # Example 2 - The following example sends diagnostic request "DefaultSession_Start"
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(canoe_cfg=r'C:\Users\Public\Documents\Vector\CANoe\Sample Configurations 11.0.81\.\CAN\Diagnostics\UDSBasic\UDSBasic.cfg')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> resp = canoe_inst.send_diag_request('Door', 'DefaultSession_Start', False)
            >>> print(resp)
            >>> canoe_inst.stop_measurement()
        """
        diag_response_data = ""
        diag_response_including_sender_name = {}
        try:
            if diag_ecu_qualifier_name in self.__diag_devices.keys():
                self.log.info(f'{diag_ecu_qualifier_name}: Diagnostic Request --> {request}')
                if request_in_bytes:
                    diag_req = self.__diag_devices[diag_ecu_qualifier_name].create_request_from_stream(request)
                else:
                    diag_req = self.__diag_devices[diag_ecu_qualifier_name].create_request(request)
                diag_req.send()
                while diag_req.pending:
                    wait(0.1)
                diag_req_responses = diag_req.responses
                if len(diag_req_responses) == 0:
                    self.log.info("Diagnostic Response Not Received.")
                else:
                    for diag_res in diag_req_responses:
                        diag_response_data = diag_res.stream
                        diag_response_including_sender_name[diag_res.sender] = diag_response_data
                        if diag_res.positive:
                            self.log.info(f"{diag_res.sender}: Diagnostic Response +ve <-- {diag_response_data}")
                        else:
                            self.log.info(f"{diag_res.Sender}: Diagnostic Response -ve <-- {diag_response_data}")
            else:
                self.log.info(
                    f'Diagnostic ECU qualifier({diag_ecu_qualifier_name}) not available in loaded CANoe config.')
        except Exception as e:
            self.log.info(f'failed to send diagnostic request({request}). {e}')
        return diag_response_including_sender_name if return_sender_name else diag_response_data

    def control_tester_present(self, diag_ecu_qualifier_name: str, value: bool) -> None:
        """Starts/Stops sending autonomous/cyclical Tester Present requests to the ECU.

        Args:
            diag_ecu_qualifier_name (str): Diagnostic Node ECU Qualifier Name configured in "Diagnostic/ISO TP Configuration".
            value (bool): True - activate tester present. False - deactivate tester present.
        """
        if diag_ecu_qualifier_name in self.__diag_devices.keys():
            diag_device = self.__diag_devices[diag_ecu_qualifier_name]
            if diag_device.tester_present_status != value:
                if value:
                    diag_device.start_tester_present()
                    self.log.info(f'{diag_ecu_qualifier_name}: started tester present')
                else:
                    diag_device.stop_tester_present()
                    self.log.info(f'{diag_ecu_qualifier_name}: stopped tester present')
                wait(.1)
            else:
                self.log.info(f'{diag_ecu_qualifier_name}: tester present already set to {value}')
        else:
            self.log.info(f'diag ECU qualifier "{diag_ecu_qualifier_name}" not available in configuration.')

    def set_replay_block_file(self, block_name: str, recording_file_path: str) -> None:
        r"""Method for setting CANoe replay block file.

        Args:
            block_name: CANoe replay block name
            recording_file_path: CANoe replay recording file including path.

        Examples:
            >>> # The following example sets replay block file
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.set_replay_block_file(block_name='replay block name', recording_file_path='replay file including path')
            >>> canoe_inst.start_measurement()
        """
        replay_blocks = self.__replay_blocks
        if block_name in replay_blocks.keys():
            replay_block = replay_blocks[block_name]
            replay_block.path = recording_file_path
            self.log.info(f'Replay block "{block_name}" updated with "{recording_file_path}" path.')
        else:
            self.log.warning(f'Replay block "{block_name}" not available.')

    def control_replay_block(self, block_name: str, start_stop: bool) -> None:
        r"""Method for setting CANoe replay block file.

        Args:
            block_name (str): CANoe replay block name
            start_stop (bool): True to start replay block. False to Stop.

        Examples:
            >>> # The following example starts replay block
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.set_replay_block_file(block_name='replay block name', recording_file_path='replay file including path')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.control_replay_block('replay block name', True)
        """
        replay_blocks = self.__replay_blocks
        if block_name in replay_blocks.keys():
            replay_block = replay_blocks[block_name]
            if start_stop:
                replay_block.start()
            else:
                replay_block.stop()
            self.log.info(f'Replay block "{block_name}" {"Started" if start_stop else "Stopped"}.')
        else:
            self.log.warning(f'Replay block "{block_name}" not available.')

    def compile_all_capl_nodes(self) -> dict:
        r"""compiles all CAPL, XML and .NET nodes.
        """
        capl_obj = self.application.capl
        capl_obj.compile()
        wait(1)
        compile_result = capl_obj.compile_result()
        self.log.info(f'compiled all CAPL nodes successfully. result={compile_result["result"]}')
        return compile_result

    def call_capl_function(self, name: str, *arguments) -> bool:
        r"""Calls a CAPL function.
        Please note that the number of parameters must agree with that of the CAPL function.
        not possible to read return value of CAPL function at the moment. only execution status is returned.

        Args:
            name (str): The name of the CAPL function. Please make sure this name is already passed as argument during CANoe instance creation. see example for more info.
            arguments (tuple): Function parameters p1…p10 (optional).

        Returns:
            bool: CAPL function execution status. True-success, False-failed.

        Examples:
            >>> # The following example starts replay block
            >>> canoe_inst = CANoe(user_capl_functions=('addition_function', ))
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.call_capl_function('addition_function', 100, 200)
            >>> canoe_inst.stop_measurement()
        """
        capl_obj = self.application.capl
        exec_sts = capl_obj.call_capl_function(self.application.measurement.user_capl_function_obj_dict[name], *arguments)
        self.log.info(f'triggered capl function({name}). execution status = {exec_sts}.')
        return exec_sts

    def get_test_environments(self) -> dict:
        """returns dictionary of test environment names and class.
        """
        return self.__test_environments

    def get_test_modules(self, test_env_name: str) -> dict:
        """returns dictionary of test module names and class.
        """
        test_environments = self.get_test_environments()
        if len(test_environments) > 0:
            if test_env_name in test_environments.keys():
                return test_environments[test_env_name].get_all_test_modules()
            else:
                self.log.info(f'"{test_env_name}" not found in configuration.')
                return {}
        else:
            self.log.info(f'Zero test environments found in configuration. Not possible to fetch test modules.')
            return {}

    def execute_test_module(self, test_module_name: str) -> int:
        """use this method to execute test module.

        Args:
            test_module_name (str): test module name. avoid duplicate test module names in CANoe configuration.

        Returns:
            int: test module execution verdict. 0 ='VerdictNotAvailable', 1 = 'VerdictPassed', 2 = 'VerdictFailed',
        """
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
                self.log.info(f'test module "{test_module_name}" found in "{test_env_name}"')
                tm_obj.start()
                tm_obj.wait_for_completion()
                execution_result = tm_obj.verdict
                break
            else:
                continue
        if test_module_found and (execution_result == 1):
            self.log.info(
                f'test module "{test_env_name}.{test_module_name}" executed and verdict = {test_verdict[execution_result]}.')
        elif test_module_found and (execution_result != 1):
            self.log.info(
                f'test module "{test_env_name}.{test_module_name}" executed and verdict = {test_verdict[execution_result]}.')
        else:
            self.log.info(f'test module "{test_env_name}.{test_module_name}" not found. not possible to execute.')
        return execution_result

    def stop_test_module(self, env_name: str, module_name: str):
        """stops execution of test module.
        """
        test_modules = self.get_test_modules(test_env_name=env_name)
        if test_modules:
            if module_name in test_modules.keys():
                test_modules[module_name].stop()
            else:
                self.log.info(f'test module not found in "{env_name}" test environment.')
        else:
            self.log.info(f'test modules not available in "{env_name}" test environment.')

    def execute_all_test_modules_in_test_env(self, env_name: str):
        """executes all test modules available in test environment.
        """
        test_modules = self.get_test_modules(test_env_name=env_name)
        if test_modules:
            for tm_name in test_modules.keys():
                self.execute_test_module(tm_name)
        else:
            self.log.info(f'test modules not available in "{env_name}" test environment.')

    def stop_all_test_modules_in_test_env(self, env_name: str):
        """stops execution of all test modules available in test environment.
        """
        test_modules = self.get_test_modules(test_env_name=env_name)
        if test_modules:
            for tm_name in test_modules.keys():
                self.stop_test_module(env_name, tm_name)
        else:
            self.log.info(f'test modules not available in "{env_name}" test environment.')

    def execute_all_test_environments(self):
        """executes all test environments available in test setup.
        """
        test_environments = self.get_test_environments()
        if len(test_environments) > 0:
            for test_env_name in test_environments.keys():
                self.log.info(f'started executing test environment "{test_env_name}"...')
                self.execute_all_test_modules_in_test_env(test_env_name)
                self.log.info(f'completed executing test environment "{test_env_name}"')
        else:
            self.log.info(f'Zero test environments found in configuration.')

    def stop_all_test_environments(self):
        """stops execution of all test environments available in test setup.
        """
        test_environments = self.get_test_environments()
        if len(test_environments) > 0:
            for test_env_name in test_environments.keys():
                self.log.info(f'stopping test environment "{test_env_name}" execution')
                self.stop_all_test_modules_in_test_env(test_env_name)
                self.log.info(f'completed stopping test environment "{test_env_name}"')
        else:
            self.log.info(f'Zero test environments found in configuration.')

    def get_environment_variable_value(self, env_var_name: str) -> Union[int, float, str, tuple, None]:
        r"""returns a environment variable value.

        Args:
            env_var_name (str): The name of the environment variable. Ex- "float_var"

        Returns:
            Environment Variable value.

        Examples:
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> env_var_val = canoe_inst.get_environment_variable_value('float_var')
            >>> print(env_var_val)
        """
        var_value = None
        try:
            variable = self.application.environment.get_variable(env_var_name)
            var_value = variable.value if variable.type != 3 else tuple(variable.value)
            self.log.info(f'environment variable({env_var_name}) value <- {var_value}')
        except Exception as e:
            self.log.info(f'failed to get environment variable({env_var_name}) value. {e}')
        return var_value
    
    def set_environment_variable_value(self, env_var_name: str, value: Union[int, float, str, tuple]) -> None:
        r"""sets a value to environment variable.

        Args:
            env_var_name (str): The name of the environment variable. Ex- "speed".
            value (Union[int, float, str, tuple]): variable value. supported CAPL environment variable data types integer, double, string and data.

        Examples:
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.set_environment_variable_value('int_var', 123)
            >>> canoe_inst.set_environment_variable_value('float_var', 111.123)
            >>> canoe_inst.set_environment_variable_value('string_var', 'this is string variable')
            >>> canoe_inst.set_environment_variable_value('data_var', (1, 2, 3, 4, 5, 6))
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
            self.log.info(f'environment variable({env_var_name}) value set to -> {converted_value}')
        except Exception as e:
            self.log.info(f'failed to set system variable({env_var_name}) value. {e}')
