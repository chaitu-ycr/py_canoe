import os
import sys
from time import sleep as wait

file_path = os.path.dirname(os.path.abspath(__file__)).replace('/', '\\')
root_path = '\\'.join(file_path.split('\\')[:-1])
sys.path.extend([file_path, fr'{root_path}\src'])

from py_canoe import CANoe
canoe_inst = CANoe()

def test_canoe_open_new_save_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    canoe_inst.quit()
    wait(1)
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    wait(1)
    canoe_inst.new(auto_save=True)
    assert canoe_inst.save_configuration_as(fr'{file_path}\demo_cfg\demo_v10.cfg', 10, 0)
    wait(2)

def test_canoe_basic_measurement_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    meas_index = canoe_inst.get_measurement_index()
    print(f'CANoe measurement index value = {meas_index}')
    assert canoe_inst.start_measurement()
    assert canoe_inst.stop_measurement()
    meas_index = canoe_inst.get_measurement_index()
    print(f'CANoe measurement index value = {meas_index}')
    canoe_inst.set_measurement_index(meas_index+1)
    meas_index = canoe_inst.get_measurement_index()
    print(f'CANoe measurement index value = {meas_index}')
    canoe_inst.get_measurement_running_status()
    canoe_inst.reset_measurement()
    assert canoe_inst.stop_measurement()

def test_diag_request_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    assert canoe_inst.start_measurement()
    wait(1)
    resp = canoe_inst.send_diag_request('Door', '10 01')
    assert canoe_inst.stop_measurement()
    assert resp == '50 01 00 00 00 00'

def test_signal_value_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    assert canoe_inst.start_measurement()
    wait(1)
    canoe_inst.set_signal_value('CAN', 1, 'LightState', 'FlashLight', 1)
    wait(1)
    canoe_inst.check_signal_online('CAN', 1, 'LightState', 'FlashLight')
    canoe_inst.check_signal_state('CAN', 1, 'LightState', 'FlashLight')
    sig_val = canoe_inst.get_signal_value('CAN', 1, 'LightState', 'FlashLight')
    assert canoe_inst.stop_measurement()
    assert sig_val == 1
    wait(2)

def test_system_variable_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    assert canoe_inst.start_measurement()
    wait(1)
    canoe_inst.set_system_variable_value('demo::level_two_1::sys_var2', 20)
    wait(1)
    sys_var_val = canoe_inst.get_system_variable_value('demo::level_two_1::sys_var2')
    assert canoe_inst.stop_measurement()
    assert sys_var_val == 20
    canoe_inst.define_system_namespace('sys_demo')
    canoe_inst.define_system_variable('sys_demo::demo', 1)
    canoe_inst.save_configuration()
    assert canoe_inst.start_measurement()
    wait(1)
    sys_var_val = canoe_inst.get_system_variable_value('sys_demo::demo')
    assert sys_var_val == 1
    assert canoe_inst.stop_measurement()
    wait(2)

def test_canoe_open_close_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    wait(1)
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    wait(1)

def test_write_window_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    wait(1)
    canoe_inst.enable_write_window_output_file(fr'{file_path}\demo_cfg\Logs\write_win.txt')
    wait(1)
    assert canoe_inst.start_measurement()
    wait(1)
    canoe_inst.write_text_in_write_window("hello from python!")
    wait(1)
    text = canoe_inst.read_text_from_write_window()
    assert canoe_inst.stop_measurement()
    canoe_inst.disable_write_window_output_file()
    assert "hello from python!" in text
    wait(2)

def test_canoe_animation_mode_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo_offline.cfg')
    canoe_inst.start_measurement_in_animation_mode()
    wait(1)
    canoe_inst.break_measurement_in_offline_mode()
    wait(1)
    canoe_inst.step_measurement_event_in_single_step()
    wait(1)
    canoe_inst.reset_measurement_in_offline_mode()
    wait(1)
    assert canoe_inst.stop_measurement()
    wait(1)
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')

def test_quit_canoe():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo_offline.cfg')
    wait(1)
    canoe_inst.quit()
    wait(5)