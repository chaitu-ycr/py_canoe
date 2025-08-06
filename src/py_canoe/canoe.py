
import os
import pythoncom
import win32com.client
import win32com.client.gencache
from typing import Union
from py_canoe.utils.common import logger, wait
from py_canoe.utils import application
from py_canoe.utils import configuration
from py_canoe.utils import measurement
from py_canoe.utils import networks


class CANoe:
    def __init__(self, py_canoe_log_dir='', user_capl_functions=tuple()):
        pythoncom.CoInitialize()
        self.bus_type = {'CAN': 1, 'J1939': 2, 'TTP': 4, 'LIN': 5, 'MOST': 6, 'Kline': 14}
        measurement.MeasurementEvents.CAPL_FUNCTION_NAMES = user_capl_functions
        self.capl_function_objects = lambda: measurement.MeasurementEvents.CAPL_FUNCTION_OBJECTS
        self.com_object = win32com.client.gencache.EnsureDispatch("CANoe.Application")
        win32com.client.WithEvents(self.com_object, application.ApplicationEvents)
        win32com.client.WithEvents(self.com_object.Measurement, measurement.MeasurementEvents)
        win32com.client.WithEvents(self.com_object.Configuration, configuration.ConfigurationEvents)

    def __del__(self):
        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            logger.error(f"Error during COM uninitialization: {e}")

    def new(self, auto_save: bool = False, prompt_user: bool = False) -> bool:
        try:
            self.com_object.New(auto_save, prompt_user)
            logger.info('New CANoe configuration created 🎉')
            return True
        except Exception as e:
            logger.error(f"😡 Error creating new CANoe configuration: {e}")
            return False

    def open(self, canoe_cfg: str, visible=True, auto_save=True, prompt_user=False, auto_stop=True, timeout=60) -> bool:
        try:
            self.com_object.Visible = visible
            if auto_stop:
                self.stop_measurement(timeout=timeout)
            self.com_object.Open(canoe_cfg, auto_save, prompt_user)
            status = application.wait_for_event_canoe_configuration_to_open(timeout)
            if status:
                self._fetch_diagnostic_devices()
            return status
        except Exception as e:
            logger.error(f"😡 Error opening CANoe configuration '{canoe_cfg}': {e}")
            return False

    def quit(self, timeout=30) -> bool:
        try:
            self.com_object.Quit()
            return application.wait_for_event_canoe_quit(timeout)
        except Exception as e:
            logger.error(f"😡 Error quitting CANoe application: {e}")
            return False

    def start_measurement(self, timeout=30) -> bool:
        try:
            if self.com_object.Measurement.Running:
                logger.info("Measurement is already running.")
                return True
            self.com_object.Measurement.Start()
            return measurement.wait_for_event_canoe_measurement_started(timeout, self.com_object)
        except Exception as e:
            logger.error(f"😡 Error starting CANoe measurement: {e}")
            return False

    def stop_measurement(self, timeout=30) -> bool:
        return self.stop_ex_measurement(timeout)

    def stop_ex_measurement(self, timeout=60) -> bool:
        try:
            if not self.com_object.Measurement.Running:
                logger.info("Measurement is already stopped.")
                return True
            self.com_object.Measurement.StopEx()
            return measurement.wait_for_event_canoe_measurement_stopped(timeout, self.com_object)
        except Exception as e:
            logger.error(f"😡 Error stopping CANoe measurement with StopEx: {e}")
            return False

    def reset_measurement(self, timeout=30) -> bool:
        try:
            if not self.stop_ex_measurement(timeout=timeout):
                logger.error("😡 Error stopping measurement during reset.")
                return False
            if not self.start_measurement(timeout=timeout):
                logger.error("😡 Error starting measurement during reset.")
                return False
            logger.info("Measurement reset 🔁 successfully.")
            return True
        except Exception as e:
            logger.error(f"😡 Exception during measurement reset: {e}")
            return False

    def get_measurement_running_status(self) -> bool:
        return self.com_object.Measurement.Running

    def start_measurement_in_animation_mode(self, animation_delay=100, timeout=30) -> bool:
        try:
            if self.com_object.Measurement.Running:
                logger.info("Measurement is already running.")
                return True
            self.com_object.Measurement.AnimationDelay = animation_delay
            self.com_object.Measurement.Animate()
            started = measurement.wait_for_event_canoe_measurement_started(timeout, self.com_object)
            if started:
                logger.info(f"Measurement started 🏃‍♂️ in Animation mode with delay: {animation_delay} ⏲️")
            return started
        except Exception as e:
            logger.error(f"😡 Error starting CANoe measurement in animation mode: {e}")
            return False

    def break_measurement_in_offline_mode(self) -> bool:
        try:
            if not self.com_object.Measurement.Running:
                logger.info("Measurement is not running, cannot break.")
                return False
            self.com_object.Measurement.Break()
            logger.info("Measurement Break applied 🫷 in Offline mode")
            return True
        except Exception as e:
            logger.error(f"😡 Error breaking CANoe measurement in offline mode: {e}")
            return False

    def reset_measurement_in_offline_mode(self) -> bool:
        try:
            self.com_object.Measurement.Reset()
            logger.info("Measurement Reset triggered 🔁 in Offline mode")
            return True
        except Exception as e:
            logger.error(f"😡 Error resetting CANoe measurement in offline mode: {e}")
            return False

    def step_measurement_event_in_single_step(self) -> bool:
        try:
            self.com_object.Measurement.Step()
            logger.info("Measurement Step triggered in Single Step 👣 mode")
            return True
        except Exception as e:
            logger.error(f"😡 Error stepping CANoe measurement in single step mode: {e}")
            return False

    def get_measurement_index(self) -> int:
        try:
            index = self.com_object.Measurement.Index
            logger.info(f"Measurement Index retrieved: {index}")
            return index
        except Exception as e:
            logger.error(f"😡 Error retrieving CANoe measurement index: {e}")
            return -1

    def set_measurement_index(self, index: int) -> bool:
        try:
            self.com_object.Measurement.Index = index
            logger.info(f"Measurement Index set to: {index}")
            return True
        except Exception as e:
            logger.error(f"😡 Error setting CANoe measurement index: {e}")
            return False

    def save_configuration(self) -> bool:
        try:
            if self.com_object.Configuration.Saved:
                logger.info("CANoe configuration is already saved.")
                return True
            self.com_object.Configuration.Save()
            logger.info("CANoe configuration saved successfully 💾")
            return True
        except Exception as e:
            logger.error(f"😡 Error saving CANoe configuration: {e}")
            return False

    def save_configuration_as(self, path: str, major: int, minor: int, prompt_user: bool = False, create_dir: bool = True) -> bool:
        try:
            if create_dir:
                dir_path = os.path.dirname(path)
                if dir_path:
                    os.makedirs(dir_path, exist_ok=True)
            self.com_object.Configuration.SaveAs(path, major, minor, prompt_user)
            logger.info(f"CANoe configuration saved as '{path}' successfully 💾")
            return True
        except Exception as e:
            logger.error(f"😡 Error saving CANoe configuration as '{path}': {e}")
            return False

    def get_can_bus_statistics(self, channel: int) -> dict:
        try:
            can_stat_obj = self.com_object.Configuration.OnlineSetup.BusStatistics.BusStatistic(self.bus_type['CAN'], channel)
            keys = [
                'BusLoad', 'ChipState', 'Error', 'ErrorTotal', 'Extended', 'ExtendedTotal',
                'ExtendedRemote', 'ExtendedRemoteTotal', 'Overload', 'OverloadTotal', 'PeakLoad',
                'RxErrorCount', 'Standard', 'StandardTotal', 'StandardRemote', 'StandardRemoteTotal',
                'TxErrorCount'
            ]
            can_bus_stat_info = {key.lower(): getattr(can_stat_obj, key) for key in keys}
            logger.info(f'📜 CAN bus channel ({channel}) statistics:')
            for key, value in can_bus_stat_info.items():
                logger.info(f"    {key}: {value}")
            return can_bus_stat_info
        except Exception as e:
            logger.error(f"😡 Error retrieving CAN bus statistics for channel {channel}: {e}")
            return {}

    def get_canoe_version_info(self) -> dict:
        try:
            version = self.com_object.Version
            version_info = {
                'full_name': getattr(version, 'FullName', None),
                'name': getattr(version, 'Name', None),
                'major': getattr(version, 'major', None),
                'minor': getattr(version, 'minor', None),
                'build': getattr(version, 'Build', None),
                'patch': getattr(version, 'Patch', None)
            }
            logger.info('📜 CANoe Version Information:')
            for key, value in version_info.items():
                logger.info(f"    {key}: {value}")
            return version_info
        except Exception as e:
            logger.error(f"😡 Error retrieving CANoe version information: {e}")
            return {}

    def get_bus_databases_info(self, bus: str = 'CAN') -> dict:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return {}
            databases_info = {}
            for db_obj in self.com_object.GetBus(bus).Databases:
                info = {
                    'full_name': getattr(db_obj, 'FullName', None),
                    'path': getattr(db_obj, 'Path', None),
                    'name': getattr(db_obj, 'Name', None),
                    'channel': getattr(db_obj, 'Channel', None),
                    'com_obj': db_obj,
                }
                databases_info[info['name']] = info
            logger.info(f'📜 {bus_type} bus databases information:')
            for db_name, db_info in databases_info.items():
                logger.info(f"    {db_name}:")
                for key, value in db_info.items():
                    logger.info(f"        {key}: {value}")
            return databases_info
        except Exception as e:
            logger.error(f"😡 Error retrieving {bus} bus databases information: {e}")
            return {}

    def get_bus_nodes_info(self, bus: str = 'CAN') -> dict:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return {}
            nodes_info = {}
            for node_obj in self.com_object.GetBus(bus).Nodes:
                info = {
                    'full_name': getattr(node_obj, 'FullName', None),
                    'path': getattr(node_obj, 'Path', None),
                    'name': getattr(node_obj, 'Name', None),
                    'active': getattr(node_obj, 'Active', None),
                    'com_obj': node_obj,
                }
                nodes_info[info['name']] = info
            logger.info(f'📜 {bus_type} bus nodes information:')
            for node_name, node_info in nodes_info.items():
                logger.info(f"    {node_name}:")
                for key, value in node_info.items():
                    logger.info(f"        {key}: {value}")
            return nodes_info
        except Exception as e:
            logger.error(f"😡 Error retrieving {bus} bus nodes information: {e}")
            return {}

    def get_signal_value(self, bus: str, channel: int, message: str, signal: str, raw_value: bool = False) -> Union[int, float, None]:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return None
            signal_obj = self.com_object.GetBus(bus).GetSignal(channel, message, signal)
            value = signal_obj.RawValue if raw_value else signal_obj.Value
            logger.info(f"Signal({bus}{channel}.{message}.{signal}) value = {value}")
            return value
        except Exception as e:
            logger.error(f"😡 Error retrieving {bus} bus signal value: {e}")
            return None

    def set_signal_value(self, bus: str, channel: int, message: str, signal: str, value: Union[int, float], raw_value: bool = False) -> bool:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return False
            signal_obj = self.com_object.GetBus(bus).GetSignal(channel, message, signal)
            setattr(signal_obj, "RawValue" if raw_value else "Value", value)
            logger.info(f"Signal({bus}{channel}.{message}.{signal}) value set to {value}")
            return True
        except Exception as e:
            logger.error(f"😡 Error setting {bus} bus signal value: {e}")
            return False

    def get_signal_full_name(self, bus: str, channel: int, message: str, signal: str) -> Union[str, None]:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return None
            signal_obj = self.com_object.GetBus(bus).GetSignal(channel, message, signal)
            full_name = getattr(signal_obj, 'FullName', None)
            logger.info(f'Signal full name = {full_name}')
            return full_name
        except Exception as e:
            logger.error(f"😡 Error retrieving {bus} bus signal full name: {e}")
            return None

    def check_signal_online(self, bus: str, channel: int, message: str, signal: str) -> bool:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return False
            signal_obj = self.com_object.GetBus(bus).GetSignal(channel, message, signal)
            is_online = signal_obj.IsOnline
            logger.info(f'Signal({bus}{channel}.{message}.{signal}) is online: {is_online}')
            return is_online
        except Exception as e:
            logger.error(f"😡 Error checking {bus} bus signal online status: {e}")
            return False

    def check_signal_state(self, bus: str, channel: int, message: str, signal: str) -> int:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return -1
            signal_obj = self.com_object.GetBus(bus).GetSignal(channel, message, signal)
            state = signal_obj.State
            logger.info(f'Signal({bus}{channel}.{message}.{signal}) state: {state}')
            return state
        except Exception as e:
            logger.error(f"😡 Error checking {bus} bus signal state: {e}")
            return -1

    def get_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, raw_value=False) -> Union[float, int, None]:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return None
            signal_obj = self.com_object.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
            signal_value = signal_obj.RawValue if raw_value else signal_obj.Value
            logger.info(f'J1939 Signal({bus}{channel}.{message}.{signal}) value = {signal_value}')
            return signal_value
        except Exception as e:
            logger.error(f"😡 Error retrieving J1939 bus signal value: {e}")
            return None

    def set_j1939_signal_value(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int, value: Union[float, int], raw_value: bool = False) -> bool:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return False
            signal_obj = self.com_object.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
            setattr(signal_obj, "RawValue" if raw_value else "Value", value)
            logger.info(f'J1939 Signal({bus}{channel}.{message}.{signal}) value set to {value}')
            return True
        except Exception as e:
            logger.error(f"😡 Error setting J1939 bus signal value: {e}")
            return False

    def get_j1939_signal_full_name(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> Union[str, None]:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return None
            signal_obj = self.com_object.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
            full_name = signal_obj.FullName
            logger.info(f'J1939 Signal full name = {full_name}')
            return full_name
        except Exception as e:
            logger.error(f"😡 Error retrieving J1939 bus signal full name: {e}")
            return None

    def check_j1939_signal_online(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> bool:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return False
            signal_obj = self.com_object.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
            is_online = signal_obj.IsOnline
            logger.info(f'J1939 Signal({bus}{channel}.{message}.{signal}) is online: {is_online}')
            return is_online
        except Exception as e:
            logger.error(f"😡 Error checking J1939 bus signal online status: {e}")
            return False

    def check_j1939_signal_state(self, bus: str, channel: int, message: str, signal: str, source_addr: int, dest_addr: int) -> int:
        try:
            bus_type = bus.upper()
            if bus_type not in self.bus_type:
                logger.error(f"😡 Invalid bus type '{bus_type}'. Supported types: {', '.join(self.bus_type)}")
                return -1
            signal_obj = self.com_object.GetBus(bus).GetJ1939Signal(channel, message, signal, source_addr, dest_addr)
            state = signal_obj.State
            logger.info(f'J1939 Signal({bus}{channel}.{message}.{signal}) state: {state}')
            return state
        except Exception as e:
            logger.error(f"😡 Error checking J1939 bus signal state: {e}")
            return -1

    def ui_activate_desktop(self, name: str) -> bool:
        try:
            self.com_object.UI.ActivateDesktop(name)
            logger.info(f"UI Desktop '{name}' activated successfully.")
            return True
        except Exception as e:
            logger.error(f"😡 Error activating UI Desktop '{name}': {e}")
            return False

    def ui_open_baudrate_dialog(self) -> bool:
        try:
            self.com_object.UI.OpenBaudrateDialog()
            logger.info("UI Baudrate Dialog opened successfully.")
            return True
        except Exception as e:
            logger.error(f"😡 Error opening UI Baudrate Dialog: {e}")
            return False

    def write_text_in_write_window(self, text: str) -> bool:
        try:
            self.com_object.UI.Write.Output(text)
            logger.info(f"Text written in Write Window: {text}")
            return True
        except Exception as e:
            logger.error(f"😡 Error writing text in Write Window: {e}")
            return False

    def read_text_from_write_window(self) -> Union[str, None]:
        try:
            text = self.com_object.UI.Write.Text
            logger.info("text read successfully from Write Window")
            for line in text.splitlines():
                logger.info(f"    {line}")
            return text
        except Exception as e:
            logger.error(f"😡 Error reading text from Write Window: {e}")
            return None

    def clear_write_window_content(self) -> bool:
        try:
            self.com_object.UI.Write.Clear()
            logger.info("Write Window content cleared successfully.")
            return True
        except Exception as e:
            logger.error(f"😡 Error clearing Write Window content: {e}")
            return False

    def copy_write_window_content(self) -> bool:
        try:
            self.com_object.UI.Write.Copy()
            logger.info("Write Window content copied to clipboard successfully.")
            return True
        except Exception as e:
            logger.error(f"😡 Error copying Write Window content: {e}")
            return False

    def enable_write_window_output_file(self, output_file: str, tab_index=None) -> bool:
        try:
            if tab_index is not None:
                self.com_object.UI.Write.EnableOutputFile(output_file, tab_index)
            else:
                self.com_object.UI.Write.EnableOutputFile(output_file)
            logger.info(f"Write Window output file enabled: {output_file}")
            return True
        except Exception as e:
            logger.error(f"😡 Error enabling Write Window output file: {e}")
            return False

    def disable_write_window_output_file(self, tab_index=None) -> bool:
        try:
            if tab_index is not None:
                self.com_object.UI.Write.DisableOutputFile(tab_index)
            else:
                self.com_object.UI.Write.DisableOutputFile()
            logger.info("Write Window output file disabled successfully.")
            return True
        except Exception as e:
            logger.error(f"😡 Error disabling Write Window output file: {e}")
            return False

    def get_system_variable_value(self, sys_var_name: str, return_symbolic_name=False) -> Union[int, float, str, None]:
        try:
            parts = sys_var_name.split('::')
            if len(parts) < 2:
                logger.error(f"😡 Invalid system variable name '{sys_var_name}'. Must be in 'namespace::variable' format.")
                return None
            namespace = '::'.join(parts[:-1])
            variable_name = parts[-1]
            namespace_obj = self.com_object.System.Namespaces(namespace)
            variable_obj = namespace_obj.Variables(variable_name)
            value = variable_obj.Value
            if return_symbolic_name:
                symbolic_value = variable_obj.GetSymbolicValueName(value)
                logger.info(f"System Variable '{sys_var_name}' symbolic value: {symbolic_value}")
                return symbolic_value
            logger.info(f"System Variable '{sys_var_name}' value: {value}")
            return value
        except Exception as e:
            logger.error(f"😡 Error retrieving System Variable '{sys_var_name}': {e}")
            return None

    def set_system_variable_value(self, sys_var_name: str, value: Union[int, float, str]) -> bool:
        try:
            parts = sys_var_name.split('::')
            if len(parts) < 2:
                logger.error(f"😡 Invalid system variable name '{sys_var_name}'. Must be in 'namespace::variable' format.")
                return False
            namespace = '::'.join(parts[:-1])
            variable_name = parts[-1]
            namespace_obj = self.com_object.System.Namespaces(namespace)
            variable_obj = namespace_obj.Variables(variable_name)
            var_type = type(variable_obj.Value)
            try:
                converted_value = var_type(value)
            except Exception:
                logger.error(f"😡 Could not convert value '{value}' to type {var_type.__name__} for '{sys_var_name}'")
                return False
            variable_obj.Value = converted_value
            logger.info(f"System Variable '{sys_var_name}' set to: {converted_value} (type: {var_type.__name__})")
            return True
        except Exception as e:
            logger.error(f"😡 Error setting System Variable '{sys_var_name}': {e}")
            return False

    def set_system_variable_array_values(self, sys_var_name: str, value: tuple, index: int = 0) -> bool:
        try:
            parts = sys_var_name.split('::')
            if len(parts) < 2:
                logger.error(f"😡 Invalid system variable name '{sys_var_name}'. Must be in 'namespace::variable' format.")
                return False
            namespace = '::'.join(parts[:-1])
            variable_name = parts[-1]
            variable_obj = self.com_object.System.Namespaces(namespace).Variables(variable_name)
            arr = list(variable_obj.Value)
            if index < 0 or index + len(value) > len(arr):
                logger.error(f"😡 Not enough space in System Variable Array '{sys_var_name}' to set values.")
                return False
            value_type = type(arr[0]) if arr else type(value[0])
            arr[index:index + len(value)] = [value_type(v) for v in value]
            variable_obj.Value = tuple(arr)
            logger.info(f"System Variable Array '{sys_var_name}' set to: {arr} (type: {value_type.__name__})")
            return True
        except Exception as e:
            logger.error(f"😡 Error setting System Variable Array '{sys_var_name}': {e}")
            return False

    def _fetch_diagnostic_devices(self):
        try:
            self._diagnostic_devices = {}
            for i in range(1, self.com_object.Networks.Count + 1):
                network = self.com_object.Networks.Item(i)
                for j in range(1, network.Devices.Count + 1):
                    device = network.Devices.Item(j)
                    try:
                        diagnostic = getattr(device, 'Diagnostic', None)
                        if diagnostic:
                            self._diagnostic_devices[device.Name] = diagnostic
                    except Exception:
                        pass
        except Exception as e:
            logger.error(f"😡 Error fetching Diagnostic Devices: {e}")
            return None

    def send_diag_request(self, diag_ecu_qualifier_name: str, request: str, request_in_bytes=True, return_sender_name=False) -> Union[str, dict]:
        try:
            diag_device = self._diagnostic_devices.get(diag_ecu_qualifier_name)
            if diag_device:
                if request_in_bytes:
                    diag_req_in_bytes = bytearray()
                    byte_stream = ''.join(request.split(' '))
                    for i in range(0, len(byte_stream), 2):
                        diag_req_in_bytes.append(int(byte_stream[i:i + 2], 16))
                    diag_request = diag_device.CreateRequestFromStream(diag_req_in_bytes)
                else:
                    diag_request = diag_device.CreateRequest(request)
                diag_request.Send()
                while diag_request.Pending:
                    wait(0.05)
                responses = [diag_request.Responses.item(i).Stream for i in range(1, diag_request.Responses.Count + 1)]
                response = " ".join(f"{d:02X}" for d in responses[0]).upper()
                return response
            else:
                logger.warning(f'⚠️ No responses received for request: {request}')
                return {"error": "No responses received"}
        except Exception as e:
            logger.error(f"😡 Error sending diagnostic request: {e}")
            return {"error": str(e)}
