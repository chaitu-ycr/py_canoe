"""Python package for controlling Vector CANoe tool"""

__version__ = "0.0.10"

# Import Python Libraries here
import os
import sys
import logging
import pythoncom
import win32com.client
from typing import Union
from logging import handlers
from time import sleep as wait


class CANoe:
    r"""The CANoe class represents the CANoe application.
    The CANoe class is the foundation for the object hierarchy.
    You can reach all other methods from the CANoe class instance.

    Examples:
        >>> # Example to open CANoe configuration, start measurement, stop measurement and close configuration.
        >>> canoe_inst = CANoe(py_canoe_log_dir=r'D:\.py_canoe')
        >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
        >>> canoe_inst.start_measurement()
        >>> wait(10)
        >>> canoe_inst.stop_measurement()
        >>> canoe_inst.quit()
    """

    def __init__(self, py_canoe_log_dir='') -> None:
        """
        Args:
            py_canoe_log_dir (str): directory to store py_canoe log. default 'D:\\.py_canoe'
        """
        self.__canoe_app_obj = None
        self.__CANOE_COM_APP_NAME = 'CANoe.Application'
        self.__BUS_TYPES = {'CAN': 1, 'J1939': 2, 'TTP': 4, 'LIN': 5, 'MOST': 6, 'Kline': 14}
        self.APP_DELAY = 1
        self.log = logging.getLogger('CANOE_LOG')
        self.__py_canoe_log_initialisation(py_canoe_log_dir)
        self.__sys_vars_obj_dictionary = {}
        self.__networks_obj_dictionary = {}
        self.__network_devices_obj_dictionary = {}
        self.__diag_ecu_qualifiers_dictionary = {}
        self.__replay_blocks_obj_dictionary = {}
        self.__simulation_nodes_obj_dictionary = {}

    def __py_canoe_log_initialisation(self, py_canoe_log_dir):
        self.log.setLevel(logging.DEBUG)
        log_format = logging.Formatter("%(asctime)s [CANOE_LOG] [%(levelname)-5.5s]  %(message)s")
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(log_format)
        self.log.addHandler(ch)
        if py_canoe_log_dir != '' and not os.path.exists(py_canoe_log_dir):
            os.makedirs(py_canoe_log_dir, exist_ok=True)
        if os.path.exists(py_canoe_log_dir):
            fh = handlers.RotatingFileHandler(fr'{py_canoe_log_dir}\py_canoe.log', maxBytes=(1024 * 50), backupCount=20)
            fh.setFormatter(log_format)
            self.log.addHandler(fh)

    def __dispatch_canoe(self) -> None:
        if self.__canoe_app_obj is None:
            pythoncom.CoInitialize()
            self.__canoe_app_obj = win32com.client.Dispatch(self.__CANOE_COM_APP_NAME)
            self.log.info('Dispatched CANoe win32com client.')
        else:
            self.log.info('CANoe win32com client already Dispatched')

    def __fetch_canoe_cfg_general_data(self):
        system_namespaces_obj = self.__canoe_app_obj.System.Namespaces
        self.__ui_obj = self.__canoe_app_obj.UI
        # self.__version_obj = self.__canoe_app_obj.Version

        def fetch_variables(namespace_obj, namespace_name):
            variables_obj = namespace_obj.Variables
            for variable_obj in variables_obj:
                variable_name = f"{namespace_name}::{variable_obj.Name}"
                self.__sys_vars_obj_dictionary[variable_name] = variable_obj

        def fetch_namespaces(namespace_obj, obj_name):
            fetch_variables(namespace_obj, obj_name)
            for ns in namespace_obj.Namespaces:
                fetch_namespaces(ns, f'{obj_name}::{ns.Name}')

        for namespace in system_namespaces_obj:
            fetch_namespaces(namespace, namespace.Name)
        for n in self.__canoe_app_obj.Networks:
            self.__networks_obj_dictionary[n.Name] = n
            self.__network_devices_obj_dictionary[n.Name] = {}
            for d in n.Devices:
                self.__network_devices_obj_dictionary[n.Name][d.Name] = d
                try:
                    self.__diag_ecu_qualifiers_dictionary[d.Name] = d.Diagnostic
                except pythoncom.com_error:
                    pass
        for rb in self.__canoe_app_obj.Bus.ReplayCollection:
            self.__replay_blocks_obj_dictionary[rb.Name] = rb
        for sn in self.__canoe_app_obj.Configuration.SimulationSetup.Nodes:
            self.__simulation_nodes_obj_dictionary[sn.Name] = sn

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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
        """
        if os.path.isfile(canoe_cfg):
            self.log.info(f'CANoe cfg "{canoe_cfg}" found.')
            self.__dispatch_canoe()
            self.__canoe_app_obj.Visible = visible
            self.__canoe_app_obj.Open(canoe_cfg, auto_save, prompt_user)
            self.log.info(f'loaded CANoe config "{canoe_cfg}"')
            self.__fetch_canoe_cfg_general_data()
            self.log.info('Fetched CANoe System Variables.')
        else:
            self.log.info(f'CANoe cfg "{canoe_cfg}" not found.')

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
        self.__dispatch_canoe()
        self.__canoe_app_obj.New(auto_save, prompt_user)
        self.log.info('created a new configuration')

    def quit(self) -> None:
        r"""Quits CANoe without saving changes in the configuration.
        
        Examples:
            >>> # The following example quits CANoe
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.quit()
        """
        if self.__canoe_app_obj.Measurement.Running:
            self.stop_measurement()
        self.__canoe_app_obj.Configuration.Modified = False
        self.__canoe_app_obj.Quit()
        self.log.info('CANoe Closed without saving.')

    def start_measurement_in_animation_mode(self, animation_delay=100) -> None:
        r"""Starts the measurement in Animation mode.

        Args:
            animation_delay (int): The animation delay during the measurement in Offline Mode.

        Examples:
            >>> # The following example starts the measurement in Animation mode
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement_in_animation_mode()
        """
        if not self.__canoe_app_obj.Measurement.Running:
            self.__canoe_app_obj.Measurement.AnimationDelay = animation_delay
            self.__canoe_app_obj.Measurement.Animate()
            self.log.info(f'Started the measurement in Animation mode with animation delay = {animation_delay}.')

    def break_measurement_in_offline_mode(self) -> None:
        r"""Interrupts the playback in Offline mode.

        Examples:
            >>> # The following example interrupts the playback in Offline mode
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.break_measurement_in_offline_mode()
        """
        if self.__canoe_app_obj.Measurement.Running:
            self.__canoe_app_obj.Measurement.Break()
            self.log.info('Interrupted the playback in Offline mode.')

    def reset_measurement_in_offline_mode(self) -> None:
        r"""Resets the measurement in Offline mode.

        Examples:
            >>> # The following example resets the measurement in Offline mode
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.reset_measurement_in_offline_mode()
        """
        self.__canoe_app_obj.Measurement.Reset()
        self.log.info('resetted measurement in offline mode.')

    def start_measurement(self) -> bool:
        r"""Starts the measurement.

        Returns:
            True if measurement started. else Flase.

        Examples:
            >>> # The following example starts the measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
        """
        if not self.__canoe_app_obj.Measurement.Running:
            self.__canoe_app_obj.Measurement.Start()
            if not self.__canoe_app_obj.Measurement.Running:
                self.log.info(f'waiting({self.APP_DELAY}s) for measurement to start running.')
                wait(self.APP_DELAY)
            self.log.info(f'CANoe Measurement Running Status: {self.__canoe_app_obj.Measurement.Running}')
        return self.__canoe_app_obj.Measurement.Running

    def step_measurement_event_in_single_step(self) -> None:
        r"""Processes a measurement event in single step.

        Examples:
            >>> # The following example processes a measurement event in single step
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.step_measurement_event_in_single_step()
        """
        if not self.__canoe_app_obj.Measurement.Running:
            self.__canoe_app_obj.Measurement.Step()
            self.log.info('processed a measurement event in single step')

    def stop_measurement(self) -> bool:
        r"""Stops the measurement.

        Returns:
            True if measurement stopped. else Flase.

        Examples:
            >>> # The following example stops the measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.stop_measurement()
        """
        if self.__canoe_app_obj.Measurement.Running:
            self.__canoe_app_obj.Measurement.Stop()
            for i in range(5):
                if self.__canoe_app_obj.Measurement.Running:
                    self.log.info(f'CANoe Simulation still running. waiting for {self.APP_DELAY} seconds.')
                    wait(self.APP_DELAY)
                else:
                    break
        self.log.info(f'Triggered stop measurement. Measurement running status = {self.__canoe_app_obj.Measurement.Running}')
        return not self.__canoe_app_obj.Measurement.Running

    def reset_measurement(self) -> bool:
        r"""reset the measurement.

        Returns:
            Measurement running status(True/False).

        Examples:
            >>> # The following example resets the measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.reset_measurement()
        """
        if self.__canoe_app_obj.Measurement.Running:
            self.stop_measurement()
        self.start_measurement()
        self.log.info(f'Resetted measurement. Measurement running status = {self.__canoe_app_obj.Measurement.Running}')
        return self.__canoe_app_obj.Measurement.Running

    def stop_ex_measurement(self) -> None:
        r"""StopEx repairs differences in the behavior of the Stop method on deferred stops concerning simulated and real mode in CANoe.

        Examples:
            >>> # The following example full stops the measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.stop_ex_measurement()
        """
        if self.__canoe_app_obj.Measurement.Running:
            self.__canoe_app_obj.Measurement.StopEx()
            self.log.info(f'Stopped measurement. Measurement running status = {self.__canoe_app_obj.Measurement.Running}')

    def get_measurement_index(self) -> int:
        r"""gets the measurement index for the next measurement.

        Returns:
            Measurement Index.

        Examples:
            >>> # The following example gets the measurement index measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.stop_measurement()
            >>> canoe_inst.get_measurement_index()
        """
        return self.__canoe_app_obj.Measurement.MeasurementIndex

    def set_measurement_index(self, index: int) -> int:
        r"""sets the measurement index for the next measurement.

        Args:
            index (int): index value to set.

        Returns:
            Measurement Index value.

        Examples:
            >>> # The following example sets the measurement index for the next measurement to 15
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.stop_measurement()
            >>> canoe_inst.set_measurement_index(15)
        """
        self.__canoe_app_obj.Measurement.MeasurementIndex = index
        self.log.info(f'CANoe measurement index set to {index}')
        return self.__canoe_app_obj.Measurement.MeasurementIndex

    def get_measurement_running_status(self) -> bool:
        r"""Returns the running state of the measurement.

        Returns:
            True if The measurement is running.
            False if The measurement is not running.

        Examples:
            >>> # The following example returns measurement running status (True/False)
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.get_measurement_running_status()
        """
        self.log.info(f'CANoe Measurement Running Status = {self.__canoe_app_obj.Measurement.Running}')
        return self.__canoe_app_obj.Measurement.Running

    def save_configuration(self) -> bool:
        r"""Saves the configuration.

        Returns:
            True if configuration saved. else False.

        Examples:
            >>> # The following example saves the configuration if necessary
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.save_configuration()
        """
        if not self.__canoe_app_obj.Configuration.Saved:
            self.__canoe_app_obj.Configuration.Save()
            self.log.info('CANoe Configuration saved.')
        return self.__canoe_app_obj.Configuration.Saved

    def save_configuration_as(self, path: str, major: int, minor: int) -> bool:
        r"""Saves the configuration as a different CANoe version.

        Args:
            path (str): The complete file name.
            major (int): The major version number of the target version.
            minor (int): The minor version number of the target version.

        Returns:
            True if configuration saved. else False.

        Examples:
            >>> # The following example saves the configuration as a CANoe 10.0 version
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.save_configuration_as(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo_v12.cfg', 10, 0)"""
        if not self.__canoe_app_obj.Configuration.Saved:
            self.__canoe_app_obj.Configuration.Save()
        self.__canoe_app_obj.Configuration.SaveAs(path, major, minor)
        self.log.info(f'CANoe Configuration saved as {path}.')
        return self.__canoe_app_obj.Configuration.Saved

    def get_signal_value(self, bus: str, channel: int, message: str, signal: str, raw_value=False) -> Union[float, int]:
        r"""get_signal_value Returns a Signal value.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet)(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the signal is sent.
            channel (int): The channel on which the signal is sent.
            message (str): The name of the message to which the signal belongs.
            signal (str): The name of the signal.
            raw_value (bool): return raw value of the signal if true. Default(False) is physical value.

        Returns:
            signal vaue.

        Examples:
            >>> # The following example gets signal value
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> sig_val = canoe_inst.get_signal_value('CAN', 1, 'LightState', 'FlashLight')
            >>> print(sig_val)
        """
        signal_obj = self.__canoe_app_obj.GetBus(bus).GetSignal(channel, message, signal)
        signal_value = signal_obj.RawValue if raw_value else signal_obj.Value
        self.log.info(f'value of signal({bus}{channel}.{message}.{signal})={signal_value}.')
        return signal_value

    def set_signal_value(self, bus: str, channel: int, message: str, signal: str, value: Union[float, int], raw_value=False) -> None:
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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.set_signal_value('CAN', 1, 'LightState', 'FlashLight', 1)
        """
        signal_obj = self.__canoe_app_obj.GetBus(bus).GetSignal(channel, message, signal)
        if raw_value:
            signal_obj.RawValue = value
        else:
            signal_obj.Value = value
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) value set to {value}.')

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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.check_signal_online('CAN', 1, 'LightState', 'FlashLight')
        """
        sig_online_status = self.__canoe_app_obj.GetBus(bus).GetSignal(channel, message, signal).IsOnline
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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.check_signal_state('CAN', 1, 'LightState', 'FlashLight')
        """
        sig_state = self.__canoe_app_obj.GetBus(bus).GetSignal(channel, message, signal).State
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) state = {sig_state}.')
        return sig_state

    def get_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int,
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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> sig_val = canoe_inst.get_j1939_signal_value('CAN', 1, 'LightState', 'FlashLight', 0, 1)
            >>> print(sig_val)
        """
        signal_obj = self.__canoe_app_obj.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
        signal_value = signal_obj.RawValue if raw_value else signal_obj.Value
        self.log.info(f'value of signal({bus}{channel}.{message}.{signal})={signal_value}.')
        return signal_value

    def set_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, value: Union[float, int],
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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.set_j1939_signal_value('CAN', 1, 'LightState', 'FlashLight', 0, 1, 1)
        """
        signal_obj = self.__canoe_app_obj.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
        if raw_value:
            signal_obj.RawValue = value
        else:
            signal_obj.Value = value
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) value set to {value}.')

    def get_system_variable_value(self, sys_var_name: str) -> Union[int, float, str]:
        r"""get_system_variable_value Returns a system variable value.

        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed"

        Returns:
            System Variable value.

        Examples:
            >>> # The following example gets system variable value
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> sys_var_val = canoe_inst.get_system_variable_value('sys_var_demo::speed')
            >>>print(sys_var_val)
        """
        variable_value = None
        if sys_var_name in self.__sys_vars_obj_dictionary.keys():
            variable_value = self.__sys_vars_obj_dictionary[sys_var_name].Value
            self.log.info(f'system variable({sys_var_name}) value = {variable_value}.')
        else:
            self.log.warning(f'system variable({sys_var_name}) not available in loaded CANoe config.')
        return variable_value

    def set_system_variable_value(self, sys_var_name: str, value: Union[int, float, str]) -> None:
        r"""set_system_variable_value sets a value to system variable.

        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed"
            value (Union[int, float, str]): variable value.

        Examples:
            >>> # The following example sets system variable value to 1
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.set_system_variable_value('sys_var_demo::speed', 1)
        """
        if sys_var_name in self.__sys_vars_obj_dictionary.keys():
            self.__sys_vars_obj_dictionary[sys_var_name].Value = value
            self.log.info(f'system variable({sys_var_name}) value set to {value}.')
        else:
            self.log.warning(f'system variable({sys_var_name}) not available in loaded CANoe config.')

    def send_diag_request(self, diag_ecu_qualifier_name: str, request: str, request_in_bytes=True) -> str:
        r"""The send_diag_request method represents the query of a diagnostic tester (client) to an ECU (server) in CANoe.

        Args:
            diag_ecu_qualifier_name (str): Diagnostic Node ECU Qualifier Name configured in "Diagnostic/ISO TP Configuration".
            request (str): Diagnostic request in bytes or diagnostic node qualifier name.
            request_in_bytes: True if Diagnostic request is bytes. False if you are using Qualifier name. Default is True.

        Returns:
            diagnostic response stream. Ex- "50 01 00 00 00 00"

        Examples:
            >>> # Example 1 - The following example sends diagnostic request "10 01"
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
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
        if diag_ecu_qualifier_name in self.__diag_ecu_qualifiers_dictionary.keys():
            self.log.info(f'Diag Req --> {request}')
            if request_in_bytes:
                diag_req_in_bytes = bytearray()
                request = ''.join(request.split(' '))
                for i in range(0, len(request), 2):
                    diag_req_in_bytes.append(int(request[i:i + 2], 16))
                diag_req = self.__diag_ecu_qualifiers_dictionary[diag_ecu_qualifier_name].CreateRequestFromStream(diag_req_in_bytes)
            else:
                diag_req = self.__diag_ecu_qualifiers_dictionary[diag_ecu_qualifier_name].CreateRequest(request)
            diag_req.Send()
            while diag_req.Pending:
                wait(0.1)
            if diag_req.Responses.Count == 0:
                self.log.info("Diagnostic Response Not Received.")
            else:
                for k in range(1, diag_req.Responses.Count + 1):
                    diag_res = diag_req.Responses(k)
                    if diag_res.Positive:
                        self.log.info(f"+ve response received.")
                    else:
                        self.log.info(f"-ve response received.")
                    diag_response_data = " ".join(f"{d:02X}" for d in diag_res.Stream).upper()
                self.log.info(f'Diag Res --> {diag_response_data}')
        else:
            self.log.info(f'Diag ECU qualifier({diag_ecu_qualifier_name}) not available in loaded CANoe config.')
        return diag_response_data

    def ui_activate_desktop(self, name: str) -> None:
        r"""Activates the desktop with the given name.

        Args:
            name (str): The name of the desktop to be activated.

        Examples:
            >>> # The following example switches to the desktop with the name "Configuration"
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.ui_activate_desktop("Configuration")
        """
        self.__canoe_app_obj.UI.ActivateDesktop(name)
        self.log.info(f'Activated / switched to "{name}" Desktop')

    def ui_open_baudrate_dialog(self) -> None:
        r"""opens the dialog for configuring the bus parameters. Make sure Measurement stopped when using this method.

        Examples:
            >>> # The following example opens the dialog for configuring the bus parameters
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.stop_measurement()
            >>> canoe_inst.ui_open_baudrate_dialog()
        """
        self.log.info('opened the dialog for configuring the bus parameters')
        self.__canoe_app_obj.UI.OpenBaudrateDialog()

    def write_text_in_write_window(self, text: str) -> None:
        r"""Outputs a line of text in the Write Window.
        Args:
            text (str): The text.

        Examples:
            >>> # The following example Outputs a line of text in the Write Window.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> canoe_inst.write_text_in_write_window("hello from python!")
            >>> wait(1)
            >>> print(canoe_inst.read_text_from_write_window())
        """
        self.__canoe_app_obj.UI.Write.Output(text)
        self.log.info(f'written "{text}" to Write Window')

    def read_text_from_write_window(self) -> str:
        r"""read the text contents from Write Window.

        Returns:
            The text content.

        Examples:
            >>> # The following example reads text from Write Window.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> canoe_inst.write_text_in_write_window("hello from python!")
            >>> wait(1)
            >>> print(canoe_inst.read_text_from_write_window())
        """
        return self.__canoe_app_obj.UI.Write.Text

    def clear_write_window_content(self) -> None:
        r"""Clears the contents of the Write Window.

        Examples:
            >>> # The following example clears content from Write Window.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> canoe_inst.write_text_in_write_window("hello from python!")
            >>> wait(1)
            >>> canoe_inst.clear_write_window_content()
        """
        self.__canoe_app_obj.UI.Write.Clear()
        self.log.info(f'Cleared Write Window Content.')

    def enable_write_window_output_file(self, output_file: str) -> None:
        r"""Enables logging of all outputs of the Write Window in the output file.

        Args:
            output_file (str): The complete path of the output file.

        Examples:
            >>> # The following example Enables logging of all outputs of the Write Window in the output file.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.enable_write_window_output_file(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\write_out.txt')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> canoe_inst.write_text_in_write_window("hello from python!")
            >>> wait(1)
            >>> canoe_inst.stop_measurement()
        """
        self.__canoe_app_obj.UI.Write.EnableOutputFile(output_file)
        self.log.info(f'Enabled Write Window logging. file path --> {output_file}')

    def disable_write_window_output_file(self) -> None:
        r"""Disables logging of all outputs of the Write Window.

        Examples:
            >>> # The following example Disables logging of all outputs of the Write Window.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.disable_write_window_output_file()
        """
        self.__canoe_app_obj.UI.Write.DisableOutputFile()
        self.log.info(f'Enabled Write Window logging.')

    def set_replay_block_file(self, block_name: str, recording_file_path: str) -> None:
        r"""Method for setting CANoe replay block file.

        Args:
            block_name: CANoe replay block name
            recording_file_path: CANoe replay recording file including path.

        Examples:
            >>> # The following example sets replay block file
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.set_replay_block_file(block_name='replay block name', recording_file_path='replay file including path')
            >>> canoe_inst.start_measurement()
        """
        if block_name in self.__replay_blocks_obj_dictionary.keys():
            self.__replay_blocks_obj_dictionary[block_name].Path = recording_file_path
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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.set_replay_block_file(block_name='replay block name', recording_file_path='replay file including path')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.control_replay_block('replay block name', True)
        """
        if block_name in self.__replay_blocks_obj_dictionary.keys():
            if start_stop:
                self.__replay_blocks_obj_dictionary[block_name].Start()
            else:
                self.__replay_blocks_obj_dictionary[block_name].Stop()
            self.log.info(f'Replay block "{block_name}" {"Started" if start_stop else "Stopped"}.')
        else:
            self.log.warning(f'Replay block "{block_name}" not available.')

    def get_can_bus_statistics(self, channel: int) -> dict:
        r"""Returns CAN Bus Statistics.

        Args:
            channel (int): The channel of the statistic that is to be returned.

        Returns:
            CAN bus statistics.

        Examples:
            >>> # The following example prints CAN channel 1 statistics
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> print(canoe_inst.get_can_bus_statistics(channel=1))
        """
        bus_statistics_obj = self.__canoe_app_obj.Configuration.OnlineSetup.BusStatistics.BusStatistic(self.__BUS_TYPES['CAN'], channel)
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
        return statistics_info

    def get_canoe_configuration_details(self) -> dict:
        r"""Returns Loaded CANoe configuration details.

        Returns:
            Returns Loaded CANoe configuration details.

        Examples:
            >>> # The following example returns CANoe application version relevant information.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_version_info = canoe_inst.get_canoe_configuration_details()
            >>> print(canoe_version_info)
        """
        configuration_details = {
            'canoe_app_full_name': self.__canoe_app_obj.Application.Version.FullName,
            'canoe_app_full_name_with_sp': self.__canoe_app_obj.Application.Version.Name,
            # The complete path to the currently loaded configuration
            'canoe_cfg': self.__canoe_app_obj.Configuration.FullName,
            # CANoe Mode(online/offline)
            'canoe_mode': 'online' if self.__canoe_app_obj.Configuration.Mode == 0 else 'offline',
            # Configuration ReadOnly ?
            'cfg_read_only': self.__canoe_app_obj.Configuration.ReadOnly,
            # CANoe configuration Networks count and Names List
            'networks_count': len(self.__networks_obj_dictionary.keys()),
            'networks_list': list(self.__networks_obj_dictionary.keys()),
            # CANoe Simulation Setup Nodes count and Names List
            'simulation_setup_nodes_count': len(self.__simulation_nodes_obj_dictionary.keys()),
            'simulation_setup_nodes_list': list(self.__simulation_nodes_obj_dictionary.keys()),
            # CANoe Replay Blocks count and Names List
            'simulation_setup_replay_blocks_count': len(self.__replay_blocks_obj_dictionary.keys()),
            'simulation_setup_replay_blocks_list': list(self.__replay_blocks_obj_dictionary.keys()),
            # The number of buses count
            'simulation_setup_buses_count': self.__canoe_app_obj.Configuration.SimulationSetup.Buses.Count,
            # The number of generators contained
            'simulation_setup_generators_count': self.__canoe_app_obj.Configuration.SimulationSetup.Generators.Count,
            # The number of interactive generators contained
            'simulation_setup_interactive_generators_count': self.__canoe_app_obj.Configuration.SimulationSetup.InteractiveGenerators.Count,
        }
        self.log.info('> CANoe Configuration Details <'.center(100, '='))
        for k, v in configuration_details.items():
            self.log.info(f'{k:<50}: {v}')
        self.log.info(''.center(100, '='))
        return configuration_details

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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_version_info = canoe_inst.get_canoe_version_info()
            >>> print(canoe_version_info)
        """
        version_info = {'full_name': self.__canoe_app_obj.Application.Version.FullName,
                        'name': self.__canoe_app_obj.Application.Version.Name,
                        'build': self.__canoe_app_obj.Application.Version.Build,
                        'major': self.__canoe_app_obj.Application.Version.Major,
                        'minor': self.__canoe_app_obj.Application.Version.Minor,
                        'patch': self.__canoe_app_obj.Application.Version.Patch}
        self.log.info('> CANoe Application.Version <'.center(100, '='))
        for k, v in version_info.items():
            self.log.info(f'{k:<10}: {v}')
        self.log.info(''.center(100, '='))
        return version_info
