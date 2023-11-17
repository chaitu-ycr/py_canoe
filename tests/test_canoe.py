import os
import logging
from time import sleep as wait
from py_canoe import CANoe

file_path = os.path.dirname(os.path.abspath(__file__)).replace('/', '\\')
root_path = file_path
canoe_inst = CANoe(fr'{root_path}\.py_canoe_log', ('addition_function', 'hello_world'))
logger_inst = logging.getLogger('CANOE_LOG')


def test_open_new_quit_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=True, prompt_user=True)
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=False, auto_save=True, prompt_user=True)
    canoe_inst.new(auto_save=True, prompt_user=False)
    canoe_inst.new(auto_save=True, prompt_user=True)
    canoe_inst.quit()


def test_meas_start_stop_restart_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    assert canoe_inst.start_measurement()
    assert canoe_inst.stop_measurement()
    assert canoe_inst.start_measurement()
    assert canoe_inst.reset_measurement()
    assert canoe_inst.get_measurement_running_status()
    assert canoe_inst.stop_ex_measurement()
    assert not canoe_inst.get_measurement_running_status()


def test_meas_offline_start_stop_restart_methods():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo_offline.cfg')
    canoe_inst.add_offline_source_log_file(fr'{file_path}\demo_cfg\Logs\demo_log.blf')
    canoe_inst.start_measurement_in_animation_mode(animation_delay=200)
    wait(1)
    canoe_inst.break_measurement_in_offline_mode()
    wait(1)
    canoe_inst.step_measurement_event_in_single_step()
    wait(1)
    canoe_inst.reset_measurement_in_offline_mode()
    wait(1)
    assert canoe_inst.stop_measurement()
    wait(1)


def test_meas_index_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    canoe_inst.get_measurement_index()
    assert canoe_inst.start_measurement()
    assert canoe_inst.stop_measurement()
    meas_index_old = canoe_inst.get_measurement_index()
    canoe_inst.set_measurement_index(meas_index_old + 1)
    meas_index_new = canoe_inst.get_measurement_index()
    assert meas_index_new == meas_index_old + 1
    canoe_inst.reset_measurement()
    assert canoe_inst.stop_measurement()


def test_meas_save_saveas_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    assert canoe_inst.save_configuration()    
    canoe_inst.new(auto_save=True)
    assert canoe_inst.save_configuration_as(path=fr'{file_path}\demo_cfg\demo_v10.cfg',
                                            major=10, minor=0, create_dir=True)
    wait(1)


def test_bus_stats_canoe_ver_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    canoe_inst.get_canoe_version_info()
    assert canoe_inst.start_measurement()
    wait(2)
    canoe_inst.get_can_bus_statistics(channel=1)
    assert canoe_inst.stop_measurement()


def test_bus_signal_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    canoe_inst.get_bus_databases_info('CAN')
    canoe_inst.get_bus_nodes_info('CAN')
    assert canoe_inst.start_measurement()
    wait(1)
    canoe_inst.get_signal_full_name(bus='CAN', channel=1, message='LightState', signal='FlashLight')
    canoe_inst.get_signal_value(bus='CAN', channel=1, message='LightState', signal='FlashLight', raw_value=False)
    canoe_inst.set_signal_value(bus='CAN', channel=1, message='LightState', signal='FlashLight', value=1, raw_value=False)
    canoe_inst.set_signal_value(bus='CAN', channel=1, message='LightState', signal='FlashLight', value=1, raw_value=True)
    wait(1)
    assert canoe_inst.check_signal_online(bus='CAN', channel=1, message='LightState', signal='FlashLight')
    canoe_inst.check_signal_state(bus='CAN', channel=1, message='LightState', signal='FlashLight')
    sig_val = canoe_inst.get_signal_value(bus='CAN', channel=1, message='LightState', signal='FlashLight', raw_value=True)
    assert canoe_inst.stop_measurement()
    assert sig_val == 1


def test_ui_class_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    canoe_inst.ui_activate_desktop('Configuration')
    canoe_inst.enable_write_window_output_file(fr'{file_path}\demo_cfg\Logs\write_win.txt')
    wait(1)
    assert canoe_inst.start_measurement()
    canoe_inst.clear_write_window_content()
    wait(1)
    canoe_inst.write_text_in_write_window("hello from py_canoe!")
    wait(1)
    text = canoe_inst.read_text_from_write_window()
    assert canoe_inst.stop_measurement()
    canoe_inst.disable_write_window_output_file()
    assert "hello from py_canoe!" in text
    wait(1)


def test_system_variable_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    assert canoe_inst.start_measurement()
    wait(1)
    canoe_inst.set_system_variable_value('demo::level_two_1::sys_var2', 20)
    wait(0.1)
    sys_var_val = canoe_inst.get_system_variable_value('demo::level_two_1::sys_var2')
    canoe_inst.set_system_variable_array_values('demo::int_array_var', (00, 11, 22, 33, 44, 55, 66, 77, 88, 99))
    assert set(canoe_inst.get_system_variable_value('demo::int_array_var')) == {00, 11, 22, 33, 44, 55, 66, 77, 88, 99}
    canoe_inst.set_system_variable_array_values('demo::double_array_var', (00.0, 11.1, 22.2, 33.3, 44.4))
    assert set(canoe_inst.get_system_variable_value('demo::double_array_var')) == {00.0, 11.1, 22.2, 33.3, 44.4}
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
    canoe_inst.open(fr'{file_path}\demo_cfg\demo_diag.cfg')
    assert canoe_inst.start_measurement()
    wait(1)
    resp = canoe_inst.send_diag_request('Door', 'DefaultSession_Start', False)
    canoe_inst.control_tester_present('Door', False)
    assert resp == '50 01 00 00 00 00'
    wait(2)
    canoe_inst.control_tester_present('Door', True)
    wait(5)
    resp = canoe_inst.send_diag_request('Door', '10 02')
    assert resp == '50 02 00 00 00 00'
    canoe_inst.control_tester_present('Door', False)
    wait(2)
    resp = canoe_inst.send_diag_request('Door', '10 03', return_sender_name=True)
    assert resp['Door'] == '50 03 00 00 00 00'
    assert canoe_inst.stop_measurement()


def test_replay_block_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    assert canoe_inst.start_measurement()
    wait(1)
    canoe_inst.set_replay_block_file(block_name='DemoReplayBlock', recording_file_path=fr'{file_path}\demo_cfg\Logs\demo_log.blf')
    wait(1)
    canoe_inst.control_replay_block(block_name='DemoReplayBlock', start_stop=True)
    wait(2)
    canoe_inst.control_replay_block(block_name='DemoReplayBlock', start_stop=False)
    wait(1)
    assert canoe_inst.stop_measurement()


def test_capl_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    canoe_inst.compile_all_capl_nodes()
    assert canoe_inst.start_measurement()
    wait(1)
    assert canoe_inst.call_capl_function('addition_function', 100, 200)
    assert canoe_inst.call_capl_function('hello_world')
    assert canoe_inst.stop_measurement()


def test_test_setup_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    canoe_inst.ui_activate_desktop('TestSetup')
    assert canoe_inst.start_measurement()
    wait(1)
    canoe_inst.execute_all_test_environments()
    test_environments = canoe_inst.get_test_environments()
    logger_inst.info(f'test environments names -> {list(test_environments.keys())}')
    for te_name, _ in test_environments.items():
        test_modules = canoe_inst.get_test_modules(te_name)
        logger_inst.info(f'test modules of test env({te_name}) -> {list(test_modules.keys())}')
        canoe_inst.execute_all_test_modules_in_test_env(te_name)
    canoe_inst.execute_test_module('demo_test_node_001')
    canoe_inst.execute_test_module('demo_test_node_002')
    wait(1)
    assert canoe_inst.stop_measurement()


def test_env_var_methods():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_dev.cfg', visible=True, auto_save=False, prompt_user=False)
    assert canoe_inst.start_measurement()
    wait(1)
    canoe_inst.set_environment_variable_value('int_var', 123.12)
    canoe_inst.get_environment_variable_value('int_var')
    canoe_inst.set_environment_variable_value('float_var', 111.123)
    canoe_inst.get_environment_variable_value('float_var')
    canoe_inst.set_environment_variable_value('string_var', 'this is string variable')
    canoe_inst.get_environment_variable_value('string_var')
    canoe_inst.set_environment_variable_value('data_var', (1, 2, 3, 4, 5, 6, 7))
    canoe_inst.get_environment_variable_value('data_var')
    wait(1)
    assert canoe_inst.stop_measurement()
