import os
import logging
from time import sleep as wait
from py_canoe import CANoe

file_path = os.path.dirname(os.path.abspath(__file__)).replace('/', '\\')
root_path = file_path
canoe_inst = CANoe(fr'{root_path}\.py_canoe_log', ('addition_function', 'hello_world'))
logger_inst = logging.getLogger('CANOE_LOG')

def test_application_class_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    print(f'application name: {canoe_inst.application.name}')
    print(f'application full_name: {canoe_inst.application.full_name}')
    print(f'application path: {canoe_inst.application.path}')
    print(f'application channel_mapping_name: {canoe_inst.application.channel_mapping_name}')
    print(f'application visible: {canoe_inst.application.visible}')
    canoe_inst.get_canoe_version_info()
    canoe_inst.quit()
    wait(1)
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    canoe_inst.new(auto_save=True)
    assert canoe_inst.save_configuration_as(fr'{file_path}\demo_cfg\demo_v10.cfg', 10, 0)
    wait(2)


def test_app_measurement_class_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    canoe_inst.get_measurement_index()
    assert canoe_inst.start_measurement()
    assert canoe_inst.stop_measurement()
    meas_index = canoe_inst.get_measurement_index()
    canoe_inst.set_measurement_index(meas_index + 1)
    canoe_inst.get_measurement_index()
    canoe_inst.get_measurement_running_status()
    canoe_inst.reset_measurement()
    assert canoe_inst.stop_measurement()
    canoe_inst.open(fr'{file_path}\demo_cfg\demo_offline.cfg')
    canoe_inst.add_offline_source_log_file(fr'{file_path}\demo_cfg\Logs\demo_log.blf')
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


def test_app_bus_class_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    canoe_inst.get_bus_databases_info('CAN')
    canoe_inst.get_bus_nodes_info('CAN')
    assert canoe_inst.start_measurement()
    wait(1)
    canoe_inst.get_signal_full_name('CAN', 1, 'LightState', 'FlashLight')
    canoe_inst.set_signal_value('CAN', 1, 'LightState', 'FlashLight', 1)
    wait(1)
    canoe_inst.check_signal_online('CAN', 1, 'LightState', 'FlashLight')
    canoe_inst.check_signal_state('CAN', 1, 'LightState', 'FlashLight')
    sig_val = canoe_inst.get_signal_value('CAN', 1, 'LightState', 'FlashLight')
    assert canoe_inst.stop_measurement()
    assert sig_val == 1
    wait(1)


def test_app_ui_class_methods():
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
    wait(1)


def test_system_variable_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    assert canoe_inst.start_measurement()
    wait(1)
    canoe_inst.set_system_variable_value('demo::level_two_1::sys_var2', 20)
    wait(0.1)
    sys_var_val = canoe_inst.get_system_variable_value('demo::level_two_1::sys_var2')
    canoe_inst.set_system_variable_array_values('demo::int_array_var', (00, 11, 22, 33, 44, 55, 66, 77, 88, 99))
    assert set(canoe_inst.get_system_variable_value('demo::int_array_var')) == set((00, 11, 22, 33, 44, 55, 66, 77, 88, 99))
    canoe_inst.set_system_variable_array_values('demo::double_array_var', (00.0, 11.1, 22.2, 33.3, 44.4))
    assert set(canoe_inst.get_system_variable_value('demo::double_array_var')) == set((00.0, 11.1, 22.2, 33.3, 44.4))
    canoe_inst.set_system_variable_value('demo::string_var', 'hey hello this is string variable')
    wait(0.1)
    assert canoe_inst.get_system_variable_value('demo::string_var') == 'hey hello this is string variable'
    canoe_inst.set_system_variable_value('demo::data_var', 'hey hello this is data variable')
    wait(0.1)
    assert canoe_inst.get_system_variable_value('demo::data_var') == 'hey hello this is data variable'
    assert canoe_inst.stop_measurement()
    assert sys_var_val == 20
    canoe_inst.define_system_variable('sys_demo::demo', 1)
    canoe_inst.save_configuration()
    assert canoe_inst.start_measurement()
    wait(1)
    sys_var_val = canoe_inst.get_system_variable_value('sys_demo::demo')
    assert sys_var_val == 1
    assert canoe_inst.stop_measurement()
    wait(1)


def test_diag_request_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    assert canoe_inst.start_measurement()
    wait(1)
    resp = canoe_inst.send_diag_request('Door', 'DefaultSession_Start', False)
    assert resp == '50 01 00 00 00 00'
    resp = canoe_inst.send_diag_request('Door', '10 02')
    assert resp == '50 02 00 00 00 00'
    resp = canoe_inst.send_diag_request('Door', '10 03', return_sender_name=True)
    assert resp['Door'] == '50 03 00 00 00 00'
    assert canoe_inst.stop_measurement()


def test_capl_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    canoe_inst.compile_all_capl_nodes()
    assert canoe_inst.start_measurement()
    wait(1)
    assert canoe_inst.call_capl_function('addition_function', 100, 200)
    assert canoe_inst.call_capl_function('hello_world')
    assert canoe_inst.stop_measurement()


def test_test_module_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    assert canoe_inst.start_measurement()
    wait(1)
    canoe_inst.execute_test_module('demo_test_node_001')
    canoe_inst.execute_test_module('demo_test_node_002')
    wait(1)
    assert canoe_inst.stop_measurement()
