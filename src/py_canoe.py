"""Python package for controlling Vector CANoe tool"""

__version__ = "2.0.0"

# Import Python Libraries here
import os
import pythoncom
import win32com.client
from typing import Union
from time import sleep as wait

# import CANoe utils here
from utils.py_canoe_logger import PyCanoeLogger
from utils.application import Application
from utils.bus import Bus, Signal
from utils.capl import Capl
from utils.configuration import Configuration
from utils.measurement import Measurement
from utils.networks import Networks
from utils.system import System, Namespaces, Variables
from utils.ui import Ui
from utils.version import Version

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
    def __init__(self, py_canoe_log_dir='', user_capl_functions=tuple()):
        """
        Args:
            py_canoe_log_dir (str): directory to store py_canoe log. example 'D:\\.py_canoe'
            user_capl_functions (tuple): user defined CAPL funtions to access.
        """
        pcl = PyCanoeLogger(py_canoe_log_dir)
        self.log = pcl.log
        self.app = Application(self.log)
        self.meas = object
        self.__diag_devices = dict()
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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
        """
        pythoncom.CoInitialize()
        self.app.app_com_obj = win32com.client.Dispatch('CANoe.Application')
        cav = self.app.app_com_obj.Version
        self.log.info(f'Dispatched Vector CANoe Application {cav.major}.{cav.minor}.{cav.Build}')
        self.app.app_com_obj.Configuration.Modified = False
        self.app.visible = visible
        self.app.open(path=canoe_cfg, auto_save=auto_save, prompt_user=prompt_user)
        self.meas = Measurement(self.app, self.user_capl_function_names)
        networks_obj = Networks(self.app)
        self.__diag_devices = networks_obj.fetch_diag_devices()
    
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
        self.app.new(auto_save, prompt_user)
    
    def quit(self):
        r"""Quits CANoe without saving changes in the configuration.

        Examples:
            >>> # The following example quits CANoe
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.quit()
        """
        self.app.quit()

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
        return self.meas.start()
    
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
        return self.meas.stop()
    
    def stop_ex_measurement(self) -> bool:
        r"""StopEx repairs differences in the behavior of the Stop method on deferred stops concerning simulated and real mode in CANoe.

        Returns:
            True if measurement stopped. else Flase.

        Examples:
            >>> # The following example full stops the measurement
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.stop_ex_measurement()
        """
        return self.meas.stop_ex()
    
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
        if self.meas.running:
            self.meas.stop()
        self.meas.start()
        self.log.info(f'Resetted measurement.')
        return self.meas.running
    
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
        self.meas.animation_delay = animation_delay
        self.meas.animate()
    
    def break_measurement_in_offline_mode(self) -> None:
        r"""Interrupts the playback in Offline mode.

        Examples:
            >>> # The following example interrupts the playback in Offline mode
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.break_measurement_in_offline_mode()
        """
        self.meas.break_offline_mode()
    
    def reset_measurement_in_offline_mode(self) -> None:
        r"""Resets the measurement in Offline mode.

        Examples:
            >>> # The following example resets the measurement in Offline mode
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.reset_measurement_in_offline_mode()
        """
        self.meas.reset_offline_mode()
    
    def step_measurement_event_in_single_step(self) -> None:
        r"""Processes a measurement event in single step.

        Examples:
            >>> # The following example processes a measurement event in single step
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.step_measurement_event_in_single_step()
        """
        self.meas.step()
    
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
        self.log.info(f'measurement_index value = {self.meas.measurement_index}')
        return self.meas.measurement_index

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
        self.meas.measurement_index = index
        return self.meas.measurement_index
    
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
        self.log.info(f'CANoe Measurement Running Status = {self.meas.running}')
        return self.meas.running

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
        conf_obj = Configuration(self.app)
        return conf_obj.save()
    
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
        conf_obj = Configuration(self.app)
        config_path = '\\'.join(path.split('\\')[:-1])
        if not os.path.exists(config_path) and create_dir:
            os.makedirs(config_path, exist_ok=True)
        if os.path.exists(config_path):
            conf_obj.save_as(path, major, minor, False)
            return conf_obj.saved
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
        bus_obj = Bus(self.app, bus_type=bus)
        sig_obj = Signal(bus_obj.get_signal(channel, message, signal))
        signal_value = sig_obj.raw_value if raw_value else sig_obj.value
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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.set_signal_value('CAN', 1, 'LightState', 'FlashLight', 1)
        """
        bus_obj = Bus(self.app, bus_type=bus)
        sig_obj = Signal(bus_obj.get_signal(channel, message, signal))
        if raw_value:
            sig_obj.raw_value = value
        else:
            sig_obj.value = value
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
        bus_obj = Bus(self.app, bus_type=bus)
        sig_obj = Signal(bus_obj.get_signal(channel, message, signal))
        return sig_obj.full_name

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
        bus_obj = Bus(self.app, bus_type=bus)
        sig_obj = Signal(bus_obj.get_signal(channel, message, signal))
        sig_online_status = sig_obj.is_online
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
        bus_obj = Bus(self.app, bus_type=bus)
        sig_obj = Signal(bus_obj.get_signal(channel, message, signal))
        sig_state = sig_obj.state
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
        bus_obj = Bus(self.app, bus_type=bus)
        sig_obj = Signal(bus_obj.get_j1939_signal(channel, message, signal, source_addr, dest_addr))
        signal_value = sig_obj.raw_value if raw_value else sig_obj.value
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
        bus_obj = Bus(self.app, bus_type=bus)
        sig_obj = Signal(bus_obj.get_j1939_signal(channel, message, signal, source_addr, dest_addr))
        if raw_value:
            sig_obj.raw_value = value
        else:
            sig_obj.value = value
        self.log.info(f'signal value set to {value}.')
        self.log.info(f'signal({bus}{channel}.{message}.{signal}) value set to {value}.')
    
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
        bus_obj = Bus(self.app, bus_type=bus)
        sig_obj = Signal(bus_obj.get_j1939_signal(channel, message, signal, source_addr, dest_addr))
        return sig_obj.full_name
    
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
        bus_obj = Bus(self.app, bus_type=bus)
        sig_obj = Signal(bus_obj.get_j1939_signal(channel, message, signal, source_addr, dest_addr))
        return sig_obj.is_online

    def check_j1939_signal_state(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> int:
        """Returns the state of the signal.

        Returns:
            int: State of the signal; possible values are: 0: The default value of the signal is returned. 1: The measurement is not running; the value set by the application is returned. 3: The signal has been received in the current measurement; the current value is returned.
        """
        bus_obj = Bus(self.app, bus_type=bus)
        sig_obj = Signal(bus_obj.get_j1939_signal(channel, message, signal, source_addr, dest_addr))
        return sig_obj.state

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
        ui_obj = Ui(self.app)
        ui_obj.activate_desktop(name)

    def ui_open_baudrate_dialog(self) -> None:
        r"""opens the dialog for configuring the bus parameters. Make sure Measurement stopped when using this method.

        Examples:
            >>> # The following example opens the dialog for configuring the bus parameters
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.stop_measurement()
            >>> canoe_inst.ui_open_baudrate_dialog()
        """
        ui_obj = Ui(self.app)
        ui_obj.open_baudrate_dialog()

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
        ui_obj = Ui(self.app)
        ui_obj.send_text_to_write_window(text)

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
        ui_obj = Ui(self.app)
        return ui_obj.get_write_window_text_content()

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
        ui_obj = Ui(self.app)
        ui_obj.clear_write_window_content()

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
        ui_obj = Ui(self.app)
        ui_obj.enable_write_window_logging(output_file)

    def disable_write_window_output_file(self) -> None:
        r"""Disables logging of all outputs of the Write Window.

        Examples:
            >>> # The following example Disables logging of all outputs of the Write Window.
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.disable_write_window_output_file()
        """
        ui_obj = Ui(self.app)
        ui_obj.disable_write_window_logging()

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
        conf_obj = Configuration(self.app)
        bus_types = {'CAN': 1, 'J1939': 2, 'TTP': 4, 'LIN': 5, 'MOST': 6, 'Kline': 14}
        bus_statistics_obj = conf_obj.conf_com_obj.OnlineSetup.BusStatistics.BusStatistic(bus_types['CAN'], channel)
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
            'extended': bus_statistics_obj.ExtendedTotal,
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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_version_info = canoe_inst.get_canoe_version_info()
            >>> print(canoe_version_info)
        """
        ver_obj = Version(self.app)
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

    def define_system_variable(self, sys_var_name: str, value=0) -> object:
        r"""define_system_variable Create a system variable with an initial value
        Args:
            sys_var_name (str): The name of the system variable. Ex- "sys_var_demo::speed"
            value (Union[int, float, str]): variable value. Default value 0.
        
        Returns:
            object: The new Variable object.
        
        Examples:
            >>> # The following example gets system variable value
            >>> canoe_inst = CANoe()
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.define_system_variable('sys_demo::speed', 1)
        """
        namespace = '::'.join(sys_var_name.split('::')[:-1])
        variable_name = sys_var_name.split('::')[-1]
        new_var_obj = None
        try:
            sys_obj = System(self.app)
            namespaces_obj = Namespaces(win32com.client.Dispatch(sys_obj.sys_com_obj.Namespaces))
            new_namespace_obj = namespaces_obj.add(namespace)
            vars_obj = Variables(new_namespace_obj.Variables)
            new_var_obj = vars_obj.add(variable_name, value)
            self.log.info(f'system variable({sys_var_name}) created and value set to {value}.')
        except Exception as e:
            self.log.info(f'failed to create system variable({sys_var_name}). {e}')
        return new_var_obj

    def get_system_variable_value(self, sys_var_name: str) -> Union[int, float, str, None]:
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
        return_value = None
        try:
            sys_obj = System(self.app)
            namespace_object = sys_obj.sys_com_obj.Namespaces(namespace)
            return_value = namespace_object.Variables(variable_name).Value
            self.log.info(f'system variable({sys_var_name}) value = {return_value}.')
        except Exception as e:
            self.log.info(f'failed to get system variable({sys_var_name}) value. {e}')
        return return_value

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
        try:
            sys_obj = System(self.app)
            namespace_object = sys_obj.sys_com_obj.Namespaces(namespace)
            namespace_object.Variables(variable_name).Value = value
            self.log.info(f'system variable({sys_var_name}) value set to {value}.')
        except Exception as e:
            self.log.info(f'failed to set system variable({sys_var_name}) value. {e}')

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
        try:
            if diag_ecu_qualifier_name in self.__diag_devices.keys():
                self.log.info(f'Diag Req --> {request}')
                if request_in_bytes:
                    diag_req_in_bytes = bytearray()
                    request = ''.join(request.split(' '))
                    for i in range(0, len(request), 2):
                        diag_req_in_bytes.append(int(request[i:i + 2], 16))
                    diag_req = self.__diag_devices[diag_ecu_qualifier_name].CreateRequestFromStream(diag_req_in_bytes)
                else:
                    diag_req = self.__diag_devices[diag_ecu_qualifier_name].CreateRequest(request)
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
        except Exception as e:
            self.log.info(f'failed to send diag request({request}). {e}')
        return diag_response_data

    def __fetch_replay_blocks(self) -> dict:
        replay_blocks = dict()
        try:
            bus_obj = Bus(self.app)
            for replay_block in bus_obj.ReplayCollection:
                replay_blocks[replay_block.Name] = replay_block
        except Exception as e:
            self.log.info(f'failed to fetch replay blocks. {e}')
        return replay_blocks

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
        replay_blocks = self.__fetch_replay_blocks()
        if block_name in replay_blocks.keys():
            replay_blocks[block_name].Path = recording_file_path
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
        replay_blocks = self.__fetch_replay_blocks()
        if block_name in replay_blocks.keys():
            if start_stop:
                replay_blocks[block_name].Start()
            else:
                replay_blocks[block_name].Stop()
            self.log.info(f'Replay block "{block_name}" {"Started" if start_stop else "Stopped"}.')
        else:
            self.log.warning(f'Replay block "{block_name}" not available.')

    def compile_all_capl_nodes(self) -> None:
        """compiles all CAPL, XML and .NET nodes.
        """
        capl_obj = Capl(self.app)
        capl_obj.compile()
        self.log.info(f'compiled all nodes successfully.')
    
    def call_capl_function(self, name: str, *arguments) -> bool:
        """Calls a CAPL function.
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
            >>> canoe_inst.open(r'D:\_kms_local\vector_canoe\py_canoe\demo_cfg\demo.cfg')
            >>> canoe_inst.start_measurement()
            >>> canoe_inst.call_capl_function('addition_function', 100, 200)
            >>> canoe_inst.stop_measurement()
        """
        capl_obj = Capl(self.app)
        exec_sts = capl_obj.call_capl_function(self.meas.user_capl_function_obj_dict[name], *arguments)
        self.log.info(f'triggered capl function({name}). execution status = {exec_sts}.')
        return exec_sts
