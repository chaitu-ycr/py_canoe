# [py_canoe](https://github.com/chaitu-ycr/py_canoe)

## about package

Python 🐍 Package for controlling Vector CANoe 🛶 Tool

## important links

- py_canoe github documentation [🔗 link](https://chaitu-ycr.github.io/py_canoe/)
- pypi package [🔗 link](https://pypi.org/project/py_canoe/)
- github releases [🔗 link](https://github.com/chaitu-ycr/py_canoe/releases)
- for ideas💡/sugessions please create new discussion [here](https://github.com/chaitu-ycr/py_canoe/discussions)
- create issue or request feature [here](https://github.com/chaitu-ycr/py_canoe/issues/new/choose)
- fork [py_canoe](https://github.com/chaitu-ycr/py_canoe/fork) repo and create pull request to contribute back to this project.

## prerequisites

- [Python(>=3.9)](https://www.python.org/downloads/)
- [Vector CANoe software(>=v11)](https://www.vector.com/int/en/support-downloads/download-center/)
- [visual studio code](https://code.visualstudio.com/Download)
- Windows PC(recommended win 10 os along with 16GB RAM)

## setup and installation

create python virtual environment

```bat
python -m venv .venv
```

activate virtual environment

```bat
.venv\Scripts\activate
```

upgrade pip (optional)

```bat
python -m pip install pip --upgrade
```

Install [py_canoe](https://pypi.org/project/py_canoe/) package

```bat
pip install py_canoe --upgrade
```

---

## example use cases

### import CANoe module and create CANoe object instance

```python
from py_canoe import CANoe, wait

canoe_inst = CANoe()
```

### open CANoe, start measurement, get version info, stop measurement and close canoe configuration

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo.cfg')
canoe_inst.start_measurement()
canoe_version_info = canoe_inst.get_canoe_version_info()
canoe_inst.stop_measurement()
canoe_inst.quit()
```

### restart/reset running measurement

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo.cfg')
canoe_inst.start_measurement()
canoe_inst.reset_measurement()
canoe_inst.stop_ex_measurement()
```

### open CANoe offline config and start/break/step/reset/stop measurement in offline mode

```python
canoe_inst.open(r'tests\demo_cfg\demo_offline.cfg')
canoe_inst.add_offline_source_log_file(r'tests\demo_cfg\Logs\demo_log.blf')
canoe_inst.start_measurement_in_animation_mode(animation_delay=200)
wait(1)
canoe_inst.break_measurement_in_offline_mode()
wait(1)
canoe_inst.step_measurement_event_in_single_step()
wait(1)
canoe_inst.reset_measurement_in_offline_mode()
wait(1)
canoe_inst.stop_measurement()
wait(1)
```

### get/set CANoe measurement index

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo_dev.cfg')
meas_index_value = canoe_inst.get_measurement_index()
canoe_inst.start_measurement()
canoe_inst.stop_measurement()
meas_index_value = canoe_inst.get_measurement_index()
canoe_inst.set_measurement_index(meas_index_value + 1)
meas_index_new = canoe_inst.get_measurement_index()
canoe_inst.reset_measurement()
canoe_inst.stop_measurement()
```

### save CANoe config to a different version with different name

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo_dev.cfg')
canoe_inst.save_configuration_as(path=r'tests\demo_cfg\demo_v10.cfg', major=10, minor=0, create_dir=True)
```

### get CAN bus statistics of CAN channel 1

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo_dev.cfg')
canoe_inst.start_measurement()
wait(2)
canoe_inst.get_can_bus_statistics(channel=1)
canoe_inst.stop_measurement()
```

### get/set bus signal value, check signal state and get signal full name

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo_dev.cfg')
canoe_inst.start_measurement()
wait(1)
sig_full_name = canoe_inst.get_signal_full_name(bus='CAN', channel=1, message='LightState', signal='FlashLight')
sig_value = canoe_inst.get_signal_value(bus='CAN', channel=1, message='LightState', signal='FlashLight', raw_value=False)
canoe_inst.set_signal_value(bus='CAN', channel=1, message='LightState', signal='FlashLight', value=1, raw_value=False)
wait(1)
sig_online_state = canoe_inst.check_signal_online(bus='CAN', channel=1, message='LightState', signal='FlashLight')
sig_state = canoe_inst.check_signal_state(bus='CAN', channel=1, message='LightState', signal='FlashLight')
sig_val = canoe_inst.get_signal_value(bus='CAN', channel=1, message='LightState', signal='FlashLight', raw_value=True)
canoe_inst.stop_measurement()
```

### clear write window / read text from write window / control write window output file

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo_dev.cfg')
canoe_inst.enable_write_window_output_file(r'tests\demo_cfg\Logs\write_win.txt')
wait(1)
canoe_inst.start_measurement()
canoe_inst.clear_write_window_content()
wait(1)
canoe_inst.write_text_in_write_window("hello from py_canoe!")
wait(1)
text = canoe_inst.read_text_from_write_window()
canoe_inst.stop_measurement()
canoe_inst.disable_write_window_output_file()
wait(1)
```

### switch between CANoe desktops

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo_dev.cfg')
canoe_inst.ui_activate_desktop('Configuration')
```

### get/set system variable or define system variable

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo_dev.cfg')
canoe_inst.start_measurement()
wait(1)
canoe_inst.set_system_variable_value('demo::level_two_1::sys_var2', 20)
canoe_inst.set_system_variable_value('demo::string_var', 'hey hello this is string variable')
canoe_inst.set_system_variable_value('demo::data_var', 'hey hello this is data variable')
canoe_inst.set_system_variable_array_values('demo::int_array_var', (00, 11, 22, 33, 44, 55, 66, 77, 88, 99))
wait(0.1)
sys_var_val = canoe_inst.get_system_variable_value('demo::level_two_1::sys_var2')
sys_var_val = canoe_inst.get_system_variable_value('demo::data_var')
canoe_inst.stop_measurement()
# define system variable and use it in measurement
canoe_inst.define_system_variable('sys_demo::demo', 1)
canoe_inst.save_configuration()
canoe_inst.start_measurement()
wait(1)
sys_var_val = canoe_inst.get_system_variable_value('sys_demo::demo')
canoe_inst.stop_measurement()
```

### send diagnostic request, control tester present

```python
canoe_inst.open(r'tests\demo_cfg\demo_diag.cfg')
canoe_inst.start_measurement()
wait(1)
resp = canoe_inst.send_diag_request('Door', 'DefaultSession_Start', False)
canoe_inst.control_tester_present('Door', False)
wait(2)
canoe_inst.control_tester_present('Door', True)
wait(5)
resp = canoe_inst.send_diag_request('Door', '10 02')
canoe_inst.control_tester_present('Door', False)
wait(2)
resp = canoe_inst.send_diag_request('Door', '10 03', return_sender_name=True)
canoe_inst.stop_measurement()
```

### set replay block source file / control replay block start stop

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo_dev.cfg')
canoe_inst.start_measurement()
wait(1)
canoe_inst.set_replay_block_file(block_name='DemoReplayBlock', recording_file_path=r'tests\demo_cfg\Logs\demo_log.blf')
wait(1)
canoe_inst.control_replay_block(block_name='DemoReplayBlock', start_stop=True)
wait(2)
canoe_inst.control_replay_block(block_name='DemoReplayBlock', start_stop=False)
wait(1)
canoe_inst.stop_measurement()
```

### compile CAPL nodes and call capl function

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo_dev.cfg')
canoe_inst.compile_all_capl_nodes()
canoe_inst.start_measurement()
wait(1)
canoe_inst.call_capl_function('addition_function', 100, 200)
canoe_inst.call_capl_function('hello_world')
canoe_inst.stop_measurement()
```

### execute test setup test module / test environment

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo_dev.cfg')
canoe_inst.start_measurement()
wait(1)
canoe_inst.execute_all_test_modules_in_test_env(demo_test_environment)
canoe_inst.execute_test_module('demo_test_node_002')
wait(1)
canoe_inst.stop_measurement()
```

### get/set environment variable value

```python
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo_dev.cfg')
canoe_inst.start_measurement()
wait(1)
canoe_inst.set_environment_variable_value('int_var', 123.12)
canoe_inst.set_environment_variable_value('float_var', 111.123)
canoe_inst.set_environment_variable_value('string_var', 'this is string variable')
canoe_inst.set_environment_variable_value('data_var', (1, 2, 3, 4, 5, 6, 7))
var_value = canoe_inst.get_environment_variable_value('int_var')
var_value = canoe_inst.get_environment_variable_value('float_var')
var_value = canoe_inst.get_environment_variable_value('string_var')
var_value = canoe_inst.get_environment_variable_value('data_var')
wait(1)
canoe_inst.stop_measurement()
```
