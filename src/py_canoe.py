"""Python package for controlling Vector CANoe tool"""

__version__ = "2.0.0"

# Import Python Libraries here
import os
import pythoncom
import win32com.client
from typing import Union

# import CANoe utils here
from utils.py_canoe_logger import PyCanoeLogger
from utils.application import Application
from utils.bus import Bus
from utils.configuration import Configuration
from utils.measurement import Measurement
from utils.ui import Ui


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
    def __init__(self, py_canoe_log_dir=''):
        """
        Args:
            py_canoe_log_dir (str): directory to store py_canoe log. example 'D:\\.py_canoe'
        """
        pcl = PyCanoeLogger(py_canoe_log_dir)
        self.log = pcl.log
        self.app = Application(self.log)
        self.bus = object
        self.capl = object
        self.conf = object
        self.env = object
        self.meas = object
        self.net = object
        self.perf = object
        self.sim = object
        self.sys = object
        self.ui = object
        self.ver = object
    
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
        self.meas = Measurement(self.app)
        self.conf = Configuration(self.app)

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
        return self.conf.save()
    
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
            self.conf.save_as(path, major, minor, False)
            return self.conf.saved
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
        signal_value = bus_obj.get_signal_value(channel, message, signal, raw_value)
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
        bus_obj.set_signal_value(channel, message, signal, value, raw_value)
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
        bus_obj = Bus(self.app, bus_type=bus)
        sig_online_status = bus_obj.get_signal_is_online(channel, message, signal)
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
        sig_state = bus_obj.get_signal_state(channel, message, signal)
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
        signal_value = bus_obj.get_j1939_signal_value(channel, message, signal, source_addr, dest_addr, raw_value)
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
        bus_obj.set_j1939_signal_value(channel, message, signal, source_addr, dest_addr, value, raw_value)
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

