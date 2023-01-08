import os
import sys
from time import sleep as wait

file_path = os.path.dirname(os.path.abspath(__file__)).replace('/', '\\')
root_path = '\\'.join(file_path.split('\\')[:-1])
sys.path.extend([file_path, fr'{root_path}\src'])

from py_canoe import CANoe
canoe_inst = CANoe()

def test_get_canoe_configuration_details():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    canoe_inst.get_canoe_configuration_details()


def test_check_measurement():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    start_resp = canoe_inst.start_measurement()
    stop_resp = canoe_inst.stop_measurement()
    assert start_resp
    assert stop_resp

def test_diag_request():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    canoe_inst.start_measurement()
    wait(1)
    resp = canoe_inst.send_diag_request('Door', '10 01')
    canoe_inst.stop_measurement()
    assert resp == '50 01 00 00 00 00'

def test_write_window():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    canoe_inst.start_measurement()
    wait(1)
    canoe_inst.write_text_in_write_window("hello from python!")
    wait(1)
    text = canoe_inst.read_text_from_write_window()
    canoe_inst.stop_measurement()
    assert "hello from python!" in text

def test_set_get_signal_value():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    canoe_inst.start_measurement()
    wait(1)
    canoe_inst.set_signal_value('CAN', 1, 'LightState', 'FlashLight', 1)
    wait(1)
    sig_val = canoe_inst.get_signal_value('CAN', 1, 'LightState', 'FlashLight')
    canoe_inst.stop_measurement()
    assert sig_val == 1

def test_set_get_system_variable_value():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    canoe_inst.start_measurement()
    wait(1)
    canoe_inst.set_system_variable_value('demo::level_two_1::sys_var2', 20)
    wait(1)
    sys_var_val = canoe_inst.get_system_variable_value('demo::level_two_1::sys_var2')
    canoe_inst.stop_measurement()
    assert sys_var_val == 20
