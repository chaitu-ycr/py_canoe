"""Python package for controlling Vector CANoe tool"""

__version__ = "0.1.1"

# Import Python Libraries here
import os
import logging
from typing import Union
from logging import handlers
from win32com.client import *
from win32com.client.connect import *
from time import sleep as wait


def DoEvents():
    pythoncom.PumpWaitingMessages()
    wait(.1)


def DoEventsUntil(cond):
    while not cond():
        DoEvents()


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
    Started = False
    Stopped = False

    def __init__(self, py_canoe_log_dir='') -> None:
        """
        Args:
            py_canoe_log_dir (str): directory to store py_canoe log. example 'D:\\.py_canoe'
        """
        self.log = logging.getLogger('CANOE_LOG')
        self.__py_canoe_log_initialisation(py_canoe_log_dir)
        self.__canoe_objects = {}
        self.__dispatch_canoe()
        self.wait_for_start = lambda: DoEventsUntil(lambda: CANoe.Started)
        self.wait_for_stop = lambda: DoEventsUntil(lambda: CANoe.Stopped)
        self.__triggered_canoe_quit = False
        self.__BUS_TYPES = {'CAN': 1, 'J1939': 2, 'TTP': 4, 'LIN': 5, 'MOST': 6, 'Kline': 14}
        self.__diag_ecu_qualifiers_dictionary = {}

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

    def __dispatch_canoe(self):
        app = DispatchEx('CANoe.Application')
        app.Configuration.Modified = False
        ver = app.Version
        self.log.info(f'Dispatched CANoe Application {ver.major}.{ver.minor}.{ver.Build}...')
        self.__canoe_objects['Application'] = app
        self.__canoe_objects['Application.Configuration'] = self.__canoe_objects['Application'].Configuration
        self.__canoe_objects['Application.Measurement'] = self.__canoe_objects['Application'].Measurement
        self.__canoe_objects['Application.Measurement.Running'] = self.__canoe_objects['Application.Measurement'].Running
        self.wait_for_start = lambda: DoEventsUntil(lambda: CANoe.Started)
        self.wait_for_stop = lambda: DoEventsUntil(lambda: CANoe.Stopped)
        WithEvents(self.__canoe_objects['Application.Measurement'], CanoeMeasurementEvents)

    def __fetch_canoe_networks_data(self) -> dict:
        self.__canoe_objects['Application.Networks'] = self.__canoe_objects['Application'].Networks
        canoe_networks_dict = {}
        for network in self.__canoe_objects['Application.Networks']:
            network_name = network.Name
            canoe_networks_dict[network_name] = {}
            canoe_networks_dict[network_name]['network_obj'] = network
            # canoe_networks_dict[network_name]['BusType'] = network.BusType
            canoe_networks_dict[network_name]['Devices'] = {}
            for device in network.Devices:
                device_name = device.Name
                canoe_networks_dict[network_name]['Devices'][device_name] = {}
                canoe_networks_dict[network_name]['Devices'][device_name]['device_obj'] = device
                try:
                    canoe_networks_dict[network_name]['Devices'][device_name]['diagnostic_obj'] = device.Diagnostic
                    self.__diag_ecu_qualifiers_dictionary[device_name] = canoe_networks_dict[network_name]['Devices'][device_name]['diagnostic_obj']
                except pythoncom.com_error:
                    canoe_networks_dict[network_name]['Devices'][device_name]['diagnostic_obj'] = None
        return canoe_networks_dict

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
        if self.__triggered_canoe_quit:
            self.__dispatch_canoe()
        if os.path.isfile(canoe_cfg):
            self.log.info(f'CANoe cfg "{canoe_cfg}" found.')
            self.__canoe_objects['Application'].Visible = visible
            self.__canoe_objects['Application'].Open(canoe_cfg, auto_save, prompt_user)
            self.log.info(f'loaded CANoe config "{canoe_cfg}"')
            self.__fetch_canoe_networks_data()
        else:
            self.log.info(f'CANoe cfg "{canoe_cfg}" not found.')
        self.__triggered_canoe_quit = False

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
        self.__canoe_objects['Application'].New(auto_save, prompt_user)
        self.log.info('created a new configuration')

    def quit(self) -> None:
        r"""Quits CANoe without saving changes in the configuration.

        Examples:
            >>> # The following example quits CANoe
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.quit()
        """
        if self.__canoe_objects['Application'].Measurement.Running:
            self.stop_measurement()
        self.__canoe_objects['Application'].Configuration.Modified = False
        self.__canoe_objects['Application'].Quit()
        self.__triggered_canoe_quit = True
        self.log.info('CANoe Application Closed without saving configuration.')

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
        if not self.__canoe_objects['Application'].Measurement.Running:
            self.__canoe_objects['Application'].Measurement.AnimationDelay = animation_delay
            self.__canoe_objects['Application'].Measurement.Animate()
            self.log.info(f'Started the measurement in Animation mode with animation delay = {animation_delay}.')

    def break_measurement_in_offline_mode(self) -> None:
        r"""Interrupts the playback in Offline mode.

        Examples:
            >>> # The following example interrupts the playback in Offline mode
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.break_measurement_in_offline_mode()
        """
        if self.__canoe_objects['Application'].Measurement.Running:
            self.__canoe_objects['Application'].Measurement.Break()
            self.log.info('Interrupted the playback in Offline mode.')

    def reset_measurement_in_offline_mode(self) -> None:
        r"""Resets the measurement in Offline mode.

        Examples:
            >>> # The following example resets the measurement in Offline mode
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.reset_measurement_in_offline_mode()
        """
        self.__canoe_objects['Application'].Measurement.Reset()
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
        if not self.__canoe_objects['Application'].Measurement.Running:
            self.__canoe_objects['Application'].Measurement.Start()
            if not self.__canoe_objects['Application'].Measurement.Running:
                self.log.info(f'waiting for measurement to start...')
                self.wait_for_start()
            self.log.info(f'CANoe Measurement Started. Measurement running status = {self.__canoe_objects["Application"].Measurement.Running}')
        else:
            self.log.info(f'CANoe Measurement Already Running. Measurement running status = {self.__canoe_objects["Application"].Measurement.Running}')
        return self.__canoe_objects['Application'].Measurement.Running

    def step_measurement_event_in_single_step(self) -> None:
        r"""Processes a measurement event in single step.

        Examples:
            >>> # The following example processes a measurement event in single step
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.step_measurement_event_in_single_step()
        """
        if not self.__canoe_objects['Application'].Measurement.Running:
            self.__canoe_objects['Application'].Measurement.Step()
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
        if self.__canoe_objects['Application'].Measurement.Running:
            self.__canoe_objects['Application'].Measurement.Stop()
            self.wait_for_stop()
            self.log.info(f'CANoe Measurement Stopped. Measurement running status = {self.__canoe_objects["Application"].Measurement.Running}')
        else:
            self.log.info(f'CANoe Measurement Already Stopped. Measurement running status = {self.__canoe_objects["Application"].Measurement.Running}')
        return not self.__canoe_objects['Application'].Measurement.Running

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
        if self.__canoe_objects['Application'].Measurement.Running:
            self.stop_measurement()
        self.start_measurement()
        self.log.info(f'Resetted measurement. Measurement running status = {self.__canoe_objects["Application"].Measurement.Running}')
        return self.__canoe_objects['Application'].Measurement.Running

    def stop_ex_measurement(self) -> None:
        r"""StopEx repairs differences in the behavior of the Stop method on deferred stops concerning simulated and real mode in CANoe.

        Examples:
            >>> # The following example full stops the measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.stop_ex_measurement()
        """
        if self.__canoe_objects['Application'].Measurement.Running:
            self.__canoe_objects['Application'].Measurement.StopEx()
            self.log.info(f'Stopped measurement. Measurement running status = {self.__canoe_objects["Application"].Measurement.Running}')

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
        measurement_index = self.__canoe_objects['Application'].Measurement.MeasurementIndex
        self.log.info(f'measurement_index value = {measurement_index}')
        return measurement_index

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
        self.__canoe_objects['Application'].Measurement.MeasurementIndex = index
        self.log.info(f'CANoe measurement index set to {index}')
        return self.__canoe_objects['Application'].Measurement.MeasurementIndex

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
        self.log.info(f'CANoe Measurement Running Status = {self.__canoe_objects["Application"].Measurement.Running}')
        return self.__canoe_objects['Application'].Measurement.Running

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
        if not self.__canoe_objects['Application'].Configuration.Saved:
            self.__canoe_objects['Application'].Configuration.Save()
            self.log.info('CANoe Configuration saved.')
        else:
            self.log.info('CANoe Configuration already in saved state.')
        return self.__canoe_objects['Application'].Configuration.Saved

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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.save_configuration_as(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo_v12.cfg', 10, 0)"""
        config_path = '\\'.join(path.split('\\')[:-1])
        if not os.path.exists(config_path) and create_dir:
            os.makedirs(config_path, exist_ok=True)
        if os.path.exists(config_path):
            self.__canoe_objects['Application'].Configuration.SaveAs(path, major, minor)
            self.log.info(f'CANoe Configuration saved as {path}.')
            return self.__canoe_objects['Application'].Configuration.Saved
        else:
            self.log.info(f'tried creating {path}. but {config_path} directory not found.')
            return False

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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> sig_val = canoe_inst.get_signal_value('CAN', 1, 'LightState', 'FlashLight')
            >>> print(sig_val)
        """
        signal_obj = self.__canoe_objects['Application'].GetBus(bus).GetSignal(channel, message, signal)
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
        signal_obj = self.__canoe_objects['Application'].GetBus(bus).GetSignal(channel, message, signal)
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
        sig_online_status = self.__canoe_objects['Application'].GetBus(bus).GetSignal(channel, message, signal).IsOnline
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
        sig_state = self.__canoe_objects['Application'].GetBus(bus).GetSignal(channel, message, signal).State
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
        signal_obj = self.__canoe_objects['Application'].GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
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
        signal_obj = self.__canoe_objects['Application'].GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
        if raw_value:
            signal_obj.RawValue = value
        else:
            signal_obj.Value = value
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) value set to {value}.')

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
        self.__canoe_objects['Application'].UI.ActivateDesktop(name)
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
        self.__canoe_objects['Application'].UI.OpenBaudrateDialog()

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
        self.__canoe_objects['Application'].UI.Write.Output(text)
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
        return self.__canoe_objects['Application'].UI.Write.Text

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
        self.__canoe_objects['Application'].UI.Write.Clear()
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
        self.__canoe_objects['Application'].UI.Write.EnableOutputFile(output_file)
        self.log.info(f'Enabled Write Window logging. file path --> {output_file}')

    def disable_write_window_output_file(self) -> None:
        r"""Disables logging of all outputs of the Write Window.

        Examples:
            >>> # The following example Disables logging of all outputs of the Write Window.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.disable_write_window_output_file()
        """
        self.__canoe_objects['Application'].UI.Write.DisableOutputFile()
        self.log.info(f'Disabled Write Window logging.')

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
        bus_statistics_obj = self.__canoe_objects['Application'].Configuration.OnlineSetup.BusStatistics.BusStatistic(self.__BUS_TYPES['CAN'], channel)
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
        version_info = {'full_name': self.__canoe_objects['Application'].Application.Version.FullName,
                        'name': self.__canoe_objects['Application'].Version.Name,
                        'build': self.__canoe_objects['Application'].Version.Build,
                        'major': self.__canoe_objects['Application'].Version.major,
                        'minor': self.__canoe_objects['Application'].Version.minor,
                        'patch': self.__canoe_objects['Application'].Version.Patch}
        self.log.info('> CANoe Application.Version <'.center(100, '='))
        for k, v in version_info.items():
            self.log.info(f'{k:<10}: {v}')
        self.log.info(''.center(100, '='))
        return version_info

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
        namespace = '::'.join(sys_var_name.split('::')[:-1])
        variable_name = sys_var_name.split('::')[-1]
        namespace_object = self.__canoe_objects['Application'].System.Namespaces(namespace)
        variable_value = namespace_object.Variables(variable_name).Value
        self.log.info(f'system variable({sys_var_name}) value = {variable_value}.')
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
        namespace = '::'.join(sys_var_name.split('::')[:-1])
        variable_name = sys_var_name.split('::')[-1]
        namespace_object = self.__canoe_objects['Application'].System.Namespaces(namespace)
        namespace_object.Variables(variable_name).Value = value
        self.log.info(f'system variable({sys_var_name}) value set to {value}.')

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


class CanoeMeasurementEvents:
    """Handler for CANoe measurement events"""

    @staticmethod
    def OnStart():
        CANoe.Started = True
        CANoe.Stopped = False

    @staticmethod
    def OnStop():
        CANoe.Started = False
        CANoe.Stopped = True
