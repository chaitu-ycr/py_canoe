"""Python package for controlling Vector CANoe tool"""

__version__ = "0.0.5"

# Import Python Libraries here
import os
import sys
import logging
import pythoncom
import win32com.client
from typing import Union
from logging import handlers
from time import sleep as wait

# CANoe Logger initialisation and configuration
canoe_log = logging.getLogger('CANOE_LOG')
py_canoe_log_dir = r'D:\.py_canoe'
if not os.path.exists(py_canoe_log_dir):
    os.mkdir(py_canoe_log_dir)
canoe_log.setLevel(logging.DEBUG)
log_format = logging.Formatter("%(asctime)s [CANOE_LOG] [%(levelname)-5.5s]  %(message)s")
ch = logging.StreamHandler(sys.stdout)
ch.setFormatter(log_format)
canoe_log.addHandler(ch)
fh = handlers.RotatingFileHandler(fr'{py_canoe_log_dir}\py_canoe.log', maxBytes=(1048576 * 5), backupCount=7)
fh.setFormatter(log_format)
canoe_log.addHandler(fh)


class CANoe:
    r"""The CANoe class represents the CANoe application.
    The CANoe class is the foundation for the object hierarchy.
    You can reach all other methods from the CANoe class instance.

    Examples:
        >>> # Example to open CANoe configuration, start measurement, stop measurement and close configuration.
        >>> canoe_inst = CANoe()
        >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
        >>> canoe_inst.start_measurement()
        >>> wait(10)
        >>> canoe_inst.stop_measurement()
        >>> canoe_inst.quit()
    """

    def __init__(self) -> None:
        self.__canoe_app_obj = None
        self.CANOE_COM_APP_NAME = 'CANoe.Application'
        self.APP_DELAY = 2
        self.BUS_TYPES = {'CAN': 1, 'J1939': 2, 'TTP': 4, 'LIN': 5, 'MOST': 6, 'Kline': 14}

    def __dispatch_canoe(self) -> None:
        if self.__canoe_app_obj is None:
            pythoncom.CoInitialize()
            self.__canoe_app_obj = win32com.client.Dispatch(self.CANOE_COM_APP_NAME)
            canoe_log.info('Dispatched CANoe win32com client.')
        else:
            canoe_log.info('CANoe win32com client already Dispatched')

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
            canoe_log.info(f'CANoe cfg "{canoe_cfg}" found.')
            self.__dispatch_canoe()
            self.__canoe_app_obj.Visible = visible
            self.__canoe_app_obj.Open(canoe_cfg, auto_save, prompt_user)
            canoe_log.info(f'loaded CANoe config "{canoe_cfg}"')
        else:
            canoe_log.info(f'CANoe cfg "{canoe_cfg}" not found.')

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
        canoe_log.info('created a new configuration')

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
        canoe_log.info('CANoe Closed without saving.')

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
            canoe_log.info(f'Started the measurement in Animation mode with animation delay = {animation_delay}.')

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
            canoe_log.info('Interrupted the playback in Offline mode.')

    def reset_measurement_in_offline_mode(self) -> None:
        r"""Resets the measurement in Offline mode.

        Examples:
            >>> # The following example resets the measurement in Offline mode
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.reset_measurement_in_offline_mode()
        """
        self.__canoe_app_obj.Measurement.Reset()
        canoe_log.info('resetted measurement in offline mode.')

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
                canoe_log.info(f'waiting({self.APP_DELAY}s) for measurement to start running.')
                wait(self.APP_DELAY)
            canoe_log.info(f'CANoe Measurement Running Status: {self.__canoe_app_obj.Measurement.Running}')
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
            canoe_log.info('processed a measurement event in single step')

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
                    canoe_log.info(f'CANoe Simulation still running. waiting for {self.APP_DELAY} seconds.')
                    wait(self.APP_DELAY)
                else:
                    break
        canoe_log.info(f'Triggered stop measurement. Measurement running status = {self.__canoe_app_obj.Measurement.Running}')
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
        canoe_log.info(f'Resetted measurement. Measurement running status = {self.__canoe_app_obj.Measurement.Running}')
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
            canoe_log.info(f'Stopped measurement. Measurement running status = {self.__canoe_app_obj.Measurement.Running}')

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
        canoe_log.info(f'CANoe measurement index set to {index}')
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
        canoe_log.info(f'CANoe Measurement Running Status = {self.__canoe_app_obj.Measurement.Running}')
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
            canoe_log.info('CANoe Configuration saved.')
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
        canoe_log.info(f'CANoe Configuration saved as {path}.')
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
        canoe_log.info(f'value of signal({bus}{channel}.{message}.{signal})={signal_value}.')
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
        canoe_log.info(f'signal({bus}{channel}.{message}.{signal}) value set to {value}.')

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
        canoe_log.info(f'signal({bus}{channel}.{message}.{signal}) online status = {sig_online_status}.')
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
        canoe_log.info(f'signal({bus}{channel}.{message}.{signal}) state = {sig_state}.')
        return sig_state

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
        canoe_log.info(f'value of signal({bus}{channel}.{message}.{signal})={signal_value}.')
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
        canoe_log.info(f'signal({bus}{channel}.{message}.{signal}) value set to {value}.')

    def get_system_variable_value(self, namespace: str, variable: str) -> Union[int, float, str]:
        r"""get_system_variable_value Returns a system variable value.

        Args:
            namespace (str): The Bus on which the signal is sent.
            variable (str): The channel on which the signal is sent.

        Returns:
            System Variable value.

        Examples:
            >>> # The following example gets system variable value
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> sys_var_val = canoe_inst.get_system_variable_value('sys_var_demo', 'speed')
            >>>print(sys_var_val)
        """
        variable_value = self.__canoe_app_obj.System.Namespaces(namespace).Variables(variable).Value
        canoe_log.info(f'system variable({namespace}::{variable}) value = {variable_value}.')
        return variable_value

    def set_system_variable_value(self, namespace: str, variable: str, value: Union[int, float, str]) -> None:
        r"""set_system_variable_value sets a value to system variable.

        Args:
            namespace (str): The Bus on which the signal is sent.
            variable (str): The channel on which the signal is sent.
            value (Union[int, float, str]): variable value.

        Examples:
            >>> # The following example sets system variable value to 1
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.set_system_variable_value('sys_var_demo', 'speed', 1)
        """
        variable_obj = self.__canoe_app_obj.System.Namespaces(namespace).Variables(variable)
        variable_obj.Value = value
        canoe_log.info(f'system variable({namespace}::{variable}) value set to {value}.')

    def send_diag_request(self, bus: str, diag_node: str, request: str, request_in_bytes=True) -> str:
        r"""The send_diag_request method represents the query of a diagnostic tester (client) to an ECU (server) in CANoe.

        Args:
            bus (str): The Bus(CAN, LIN, FlexRay, MOST, AFDX, Ethernet) on which the request is sent.
            diag_node (str): Diagnostic Node Name configured in "Diagnostic/ISO TP Configuration".
            request (str): Diagnostic request in bytes or diagnostic node qualifier name.
            request_in_bytes: True if Diagnostic request is bytes. False if you are using Qualifier name. Default is True.

        Returns:
            diagnostic response stream. Ex- "50 01 00 00 00 00"

        Examples:
            >>> # Example 1 - The following example sends diagnostic request "10 01"
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(canoe_cfg=r'C:\Users\Public\Documents\Vector\CANoe\Sample Configurations 11.0.81\.\CAN\Diagnostics\UDSBasic\UDSBasic.cfg')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> resp = canoe_inst.send_diag_request('CAN', 'Door', '10 01')
            >>> print(resp)
            >>> canoe_inst.stop_measurement()
            >>> # Example 2 - The following example sends diagnostic request "DefaultSession_Start"
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(canoe_cfg=r'C:\Users\Public\Documents\Vector\CANoe\Sample Configurations 11.0.81\.\CAN\Diagnostics\UDSBasic\UDSBasic.cfg')
            >>> canoe_inst.start_measurement()
            >>> wait(1)
            >>> resp = canoe_inst.send_diag_request('CAN', 'Door', 'DefaultSession_Start', False)
            >>> print(resp)
            >>> canoe_inst.stop_measurement()
        """
        diag_response_data = ""
        canoe_log.info(f'Diag Req --> {request}')
        if request_in_bytes:
            diag_req_in_bytes = bytearray()
            request = ''.join(request.split(' '))
            for i in range(0, len(request), 2):
                diag_req_in_bytes.append(int(request[i:i + 2], 16))
            diag_req = self.__canoe_app_obj.Networks(bus).Devices(diag_node).Diagnostic.CreateRequestFromStream(diag_req_in_bytes)
        else:
            diag_req = self.__canoe_app_obj.Networks(bus).Devices(diag_node).Diagnostic.CreateRequest(request)
        diag_req.Send()
        while diag_req.Pending:
            wait(0.1)
        if diag_req.Responses.Count == 0:
            canoe_log.info("Diagnostic Response Not Received.")
        else:
            for k in range(1, diag_req.Responses.Count + 1):
                diag_res = diag_req.Responses(k)
                if diag_res.Positive:
                    canoe_log.info(f"+ve response received.")
                else:
                    canoe_log.info(f"-ve response received.")
                diag_response_data = " ".join(f"{d:02X}" for d in diag_res.Stream).upper()
            canoe_log.info(f'Diag Res --> {diag_response_data}')
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
        canoe_log.info(f'Activated / switched to "{name}" Desktop')

    def ui_open_baudrate_dialog(self) -> None:
        r"""opens the dialog for configuring the bus parameters. Make sure Measurement stopped when using this method.

        Examples:
            >>> # The following example opens the dialog for configuring the bus parameters
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.stop_measurement()
            >>> canoe_inst.ui_open_baudrate_dialog()
        """
        canoe_log.info('opened the dialog for configuring the bus parameters')
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
        canoe_log.info(f'written "{text}" to Write Window')

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
        canoe_log.info(f'Cleared Write Window Content.')

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
        canoe_log.info(f'Enabled Write Window logging. file path --> {output_file}')

    def disable_write_window_output_file(self) -> None:
        r"""Disables logging of all outputs of the Write Window.

        Examples:
            >>> # The following example Disables logging of all outputs of the Write Window.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.disable_write_window_output_file()
        """
        self.__canoe_app_obj.UI.Write.DisableOutputFile()
        canoe_log.info(f'Enabled Write Window logging.')

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
        count = self.__canoe_app_obj.Bus.ReplayCollection.Count
        for i in range(1, count + 1):
            name = self.__canoe_app_obj.Bus.ReplayCollection.Item(i).Name
            if name == block_name:
                self.__canoe_app_obj.Bus.ReplayCollection.Item(i).Path = recording_file_path
                canoe_log.info(f'Replay block "{block_name}" updated with "{recording_file_path}" path.')

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
        count = self.__canoe_app_obj.Bus.ReplayCollection.Count
        for i in range(1, count + 1):
            name = self.__canoe_app_obj.Bus.ReplayCollection.Item(i).Name
            if name == block_name:
                if start_stop:
                    self.__canoe_app_obj.Bus.ReplayCollection.Item(i).Start()
                else:
                    self.__canoe_app_obj.Bus.ReplayCollection.Item(i).Stop()
                canoe_log.info(f'Replay block "{block_name}" {"Started" if start_stop else "Stopped"}.')

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
        bus_statistics_obj = self.__canoe_app_obj.Configuration.OnlineSetup.BusStatistics.BusStatistic(self.BUS_TYPES['CAN'], channel)
        statistics_info = {
            # The bus load
            'bus_load': bus_statistics_obj.BusLoad,
            # The controller status
            'chip_state': bus_statistics_obj.ChipState,
            # The number of Error Frames per second
            'Error': bus_statistics_obj.Error,
            # The total number of Error Frames
            'ErrorTotal': bus_statistics_obj.ErrorTotal,
            # The number of messages with extended identifier per second
            'Extended': bus_statistics_obj.Extended,
            # The number of remote messages with extended identifier per second
            'ExtendedRemote': bus_statistics_obj.ExtendedRemote,
            # The total number of remote messages with extended identifier
            'ExtendedRemoteTotal': bus_statistics_obj.ExtendedRemoteTotal,
            # The number of overload frames per second
            'Overload': bus_statistics_obj.Overload,
            # The total number of overload frames
            'OverloadTotal': bus_statistics_obj.OverloadTotal,
            # The maximum bus load in 0.01 %
            'PeakLoad': bus_statistics_obj.PeakLoad,
            # Returns the current number of the Rx error counter
            'RxErrorCount': bus_statistics_obj.RxErrorCount,
            # The number of messages with standard identifier per second
            'Standard': bus_statistics_obj.Standard,
            # The total number of remote messages with standard identifier
            'StandardTotal': bus_statistics_obj.StandardTotal,
            # The number of remote messages with standard identifier per second
            'StandardRemote': bus_statistics_obj.StandardRemote,
            # The total number of remote messages with standard identifier
            'StandardRemoteTotal': bus_statistics_obj.StandardRemoteTotal,
            # The current number of the Tx error counter
            'TxErrorCount': bus_statistics_obj.TxErrorCount,
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
        # bus_obj = self.__canoe_app_obj.Bus
        # capl_obj = self.__canoe_app_obj.CAPL
        configuration_obj = self.__canoe_app_obj.Configuration
        # environment_obj = self.__canoe_app_obj.Environment
        # measurement_obj = self.__canoe_app_obj.Measurement
        networks_obj = self.__canoe_app_obj.Networks
        # performance_obj = self.__canoe_app_obj.Performance
        # simulation_obj = self.__canoe_app_obj.Simulation
        # system_obj = self.__canoe_app_obj.System
        # system_namespaces_obj = system_obj.Namespaces
        # ui_obj = self.__canoe_app_obj.UI
        # version_obj = self.__canoe_app_obj.Version

        simulation_setup_nodes_list = []
        replay_blocks_list = []
        for n in range(1, configuration_obj.SimulationSetup.Nodes.Count + 1):
            simulation_setup_nodes_list.append(configuration_obj.SimulationSetup.Nodes.Item(n).Name)
        for r in range(1, configuration_obj.SimulationSetup.ReplayCollection.Count + 1):
            replay_blocks_list.append(configuration_obj.SimulationSetup.ReplayCollection.Item(r).Name)
        networks_list = []
        diagnostic_nodes_list = []
        for n in range(1, networks_obj.Count + 1):
            networks_list.append(networks_obj.Item(n).Name)
            devices_obj = networks_obj.Item(n).Devices
            for d in range(1, devices_obj.Count):
                diagnostic_nodes_list.append(devices_obj.Item(d).Name)
        configuration_details = {
            'canoe_app_full_name': self.__canoe_app_obj.Application.Version.FullName,
            'canoe_app_full_name_with_sp': self.__canoe_app_obj.Application.Version.Name,
            # The complete path to the currently loaded configuration
            'canoe_cfg': configuration_obj.FullName,
            # CANoe Mode(online/offline)
            'canoe_mode': 'online' if configuration_obj.Mode == 0 else 'offline',
            # Configuration ReadOnly ?
            'cfg_read_only': configuration_obj.ReadOnly,
            # CANoe configuration Networks count and Names List
            'networks_count': networks_obj.Count,
            'networks_list': networks_list,
            # CANoe Simulation Setup Nodes count and Names List
            'simulation_setup_nodes_count': configuration_obj.SimulationSetup.Nodes.Count,
            'simulation_setup_nodes_list': simulation_setup_nodes_list,
            # CANoe Replay Blocks count and Names List
            'simulation_setup_replay_blocks_count': configuration_obj.SimulationSetup.ReplayCollection.Count,
            'simulation_setup_replay_blocks_list': replay_blocks_list,
            # The number of buses count
            'simulation_setup_buses_count': configuration_obj.SimulationSetup.Buses.Count,
            # The number of generators contained
            'simulation_setup_generators_count': configuration_obj.SimulationSetup.Generators.Count,
            # The number of interactive generators contained
            'simulation_setup_interactive_generators_count': configuration_obj.SimulationSetup.InteractiveGenerators.Count,
            # CANoe Diagnostic Node Names List
            'diagnostic_nodes_list': diagnostic_nodes_list,
        }
        canoe_log.info('> CANoe Configuration Details <'.center(100, '='))
        for k, v in configuration_details.items():
            canoe_log.info(f'{k:<50}: {v}')
        canoe_log.info(''.center(100, '='))
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
        canoe_log.info('> CANoe Application.Version <'.center(100, '='))
        for k, v in version_info.items():
            canoe_log.info(f'{k:<10}: {v}')
        canoe_log.info(''.center(100, '='))
        return version_info
