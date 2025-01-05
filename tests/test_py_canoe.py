import os
from py_canoe import CANoe, wait

class TestStandalonePyCanoe:
    @classmethod
    def setup_class(cls):
        cls.file_path = os.path.dirname(os.path.abspath(__file__))
        cls.root_path = cls.file_path
        cls.canoe_inst = CANoe(py_canoe_log_dir=os.path.join(cls.root_path, ".py_canoe_log"), user_capl_functions=('addition_function', 'hello_world'))
        demo_cfg_dir = os.path.join(cls.file_path, "demo_cfg")
        cls.canoe_cfg_one_ch = os.path.join(demo_cfg_dir, "demo_can_one_ch.cfg")
        cls.canoe_cfg_two_ch = os.path.join(demo_cfg_dir, "demo_can_two_ch.cfg")
        cls.canoe_cfg_gen_db_setup = os.path.join(demo_cfg_dir, "demo_conf_gen_db_setup.cfg")
        cls.canoe_cfg_dev = os.path.join(demo_cfg_dir, "demo_dev.cfg")
        cls.canoe_cfg_diag = os.path.join(demo_cfg_dir, "demo_diag.cfg")
        cls.canoe_cfg_eth_one_ch = os.path.join(demo_cfg_dir, "demo_eth_one_ch.cfg")
        cls.canoe_cfg_offline = os.path.join(demo_cfg_dir, "demo_offline.cfg")
        cls.canoe_cfg_test_setup = os.path.join(demo_cfg_dir, "demo_test_setup.cfg")
        cls.canoe_cfg_demo = os.path.join(demo_cfg_dir, "demo.cfg")
        cls.canoe_cfg_online_setup = os.path.join(demo_cfg_dir, "demo_online_setup.cfg")

    def test_open_new_quit_methods(self):
        self.canoe_inst.new(auto_save=False, prompt_user=False)
        self.canoe_inst.quit()
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=True, auto_save=False, prompt_user=False)
        self.canoe_inst.quit()
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=True, auto_save=True, prompt_user=False)
        self.canoe_inst.new(auto_save=True, prompt_user=False)
        self.canoe_inst.quit()
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=True, auto_save=True, prompt_user=True)
        self.canoe_inst.new(auto_save=True, prompt_user=True)
        self.canoe_inst.quit()
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=False, auto_save=True, prompt_user=True)
        self.canoe_inst.quit()
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=False, auto_save=False, prompt_user=True)
        self.canoe_inst.quit()
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=False, auto_save=False, prompt_user=False)
        self.canoe_inst.new(auto_save=False, prompt_user=True)
        self.canoe_inst.quit()

    def test_meas_start_stop_restart_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=True, auto_save=False, prompt_user=False)
        assert self.canoe_inst.start_measurement()
        assert self.canoe_inst.stop_measurement()
        assert self.canoe_inst.start_measurement()
        assert self.canoe_inst.reset_measurement()
        assert self.canoe_inst.get_measurement_running_status()
        assert self.canoe_inst.stop_ex_measurement()
        assert not self.canoe_inst.get_measurement_running_status()
        self.canoe_inst.quit()

    def test_meas_offline_start_stop_restart_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_offline)
        self.canoe_inst.add_offline_source_log_file(fr'{self.file_path}\demo_cfg\Logs\demo_log.blf')
        self.canoe_inst.start_measurement_in_animation_mode(animation_delay=200)
        self.canoe_inst.break_measurement_in_offline_mode()
        self.canoe_inst.step_measurement_event_in_single_step()
        self.canoe_inst.reset_measurement_in_offline_mode()
        assert self.canoe_inst.stop_measurement()

    def test_meas_index_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=True, auto_save=False, prompt_user=False)
        self.canoe_inst.get_measurement_index()
        assert self.canoe_inst.start_measurement()
        assert self.canoe_inst.stop_measurement()
        meas_index_old = self.canoe_inst.get_measurement_index()
        self.canoe_inst.set_measurement_index(meas_index_old + 1)
        meas_index_new = self.canoe_inst.get_measurement_index()
        assert meas_index_new == meas_index_old + 1
        self.canoe_inst.reset_measurement()
        assert self.canoe_inst.stop_measurement()

    def test_meas_save_saveas_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev)
        assert self.canoe_inst.save_configuration()
        assert self.canoe_inst.save_configuration_as(path=fr'{self.file_path}\demo_cfg\demo_v10.cfg',
                                                major=10, minor=0, create_dir=True)
        wait(1)

    def test_bus_stats_canoe_ver_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=True, auto_save=False, prompt_user=False)
        self.canoe_inst.get_canoe_version_info()
        assert self.canoe_inst.start_measurement()
        wait(2)
        self.canoe_inst.get_can_bus_statistics(channel=1)
        assert self.canoe_inst.stop_measurement()

    def test_bus_signal_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=True, auto_save=False, prompt_user=False)
        self.canoe_inst.get_bus_databases_info('CAN')
        self.canoe_inst.get_bus_nodes_info('CAN')
        assert self.canoe_inst.start_measurement()
        wait(5)
        self.canoe_inst.get_signal_full_name(bus='CAN', channel=1, message='LightState', signal='FlashLight')
        self.canoe_inst.get_signal_value(bus='CAN', channel=1, message='LightState', signal='FlashLight', raw_value=False)
        self.canoe_inst.set_signal_value(bus='CAN', channel=1, message='LightState', signal='FlashLight', value=1, raw_value=False)
        self.canoe_inst.set_signal_value(bus='CAN', channel=1, message='LightState', signal='FlashLight', value=1, raw_value=True)
        wait(1)
        assert self.canoe_inst.check_signal_online(bus='CAN', channel=1, message='LightState', signal='FlashLight')
        self.canoe_inst.check_signal_state(bus='CAN', channel=1, message='LightState', signal='FlashLight')
        sig_val = self.canoe_inst.get_signal_value(bus='CAN', channel=1, message='LightState', signal='FlashLight', raw_value=True)
        assert self.canoe_inst.stop_measurement()
        assert sig_val == 1

    def test_ui_class_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev)
        self.canoe_inst.ui_activate_desktop('Configuration')
        self.canoe_inst.enable_write_window_output_file(fr'{self.file_path}\demo_cfg\Logs\write_win.txt')
        wait(1)
        assert self.canoe_inst.start_measurement()
        self.canoe_inst.clear_write_window_content()
        wait(1)
        self.canoe_inst.write_text_in_write_window("hello from py_canoe!")
        wait(1)
        text = self.canoe_inst.read_text_from_write_window()
        assert self.canoe_inst.stop_measurement()
        self.canoe_inst.disable_write_window_output_file()
        assert "hello from py_canoe!" in text
        wait(1)

    def test_system_variable_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev)
        assert self.canoe_inst.start_measurement()
        wait(1)
        self.canoe_inst.set_system_variable_value('demo::level_two_1::sys_var2', 20)
        wait(0.1)
        sys_var_val = self.canoe_inst.get_system_variable_value('demo::level_two_1::sys_var2')
        self.canoe_inst.set_system_variable_array_values('demo::int_array_var', (00, 11, 22, 33, 44, 55, 66, 77, 88, 99))
        assert set(self.canoe_inst.get_system_variable_value('demo::int_array_var')) == {00, 11, 22, 33, 44, 55, 66, 77, 88, 99}
        self.canoe_inst.set_system_variable_array_values('demo::double_array_var', (00.0, 11.1, 22.2, 33.3, 44.4))
        assert set(self.canoe_inst.get_system_variable_value('demo::double_array_var')) == {00.0, 11.1, 22.2, 33.3, 44.4}
        self.canoe_inst.set_system_variable_value('demo::string_var', 'hey hello this is string variable')
        wait(0.1)
        assert self.canoe_inst.get_system_variable_value('demo::string_var') == 'hey hello this is string variable'
        self.canoe_inst.set_system_variable_value('demo::data_var', 'hey hello this is data variable')
        wait(0.1)
        assert self.canoe_inst.get_system_variable_value('demo::data_var') == 'hey hello this is data variable'
        assert self.canoe_inst.stop_measurement()
        assert sys_var_val == 20
        self.canoe_inst.define_system_variable('sys_demo::demo', 1)
        self.canoe_inst.save_configuration()
        assert self.canoe_inst.start_measurement()
        wait(1)
        sys_var_val = self.canoe_inst.get_system_variable_value('sys_demo::demo')
        assert sys_var_val == 1
        sys_var_val_name = self.canoe_inst.get_system_variable_value('demo::var_on_off', True)
        assert sys_var_val_name == 'On'
        assert self.canoe_inst.stop_measurement()
        wait(1)

    def test_diag_request_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_diag)
        assert self.canoe_inst.start_measurement()
        wait(1)
        resp = self.canoe_inst.send_diag_request('Door', 'DefaultSession_Start', False)
        self.canoe_inst.control_tester_present('Door', False)
        assert resp == '50 01 00 00 00 00'
        wait(2)
        self.canoe_inst.control_tester_present('Door', True)
        wait(5)
        resp = self.canoe_inst.send_diag_request('Door', '10 02')
        assert resp == '50 02 00 00 00 00'
        self.canoe_inst.control_tester_present('Door', False)
        wait(2)
        resp = self.canoe_inst.send_diag_request('Door', '10 03', return_sender_name=True)
        assert resp['Door'] == '50 03 00 00 00 00'
        wait(2)
        resp = self.canoe_inst.send_diag_request('Door', '22 F1 AA', return_sender_name=True)
        assert '62 F1 AA' in resp['Door']
        assert self.canoe_inst.stop_measurement()

    def test_replay_block_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=True, auto_save=True, prompt_user=False, auto_stop=True)
        assert self.canoe_inst.start_measurement()
        wait(1)
        self.canoe_inst.set_replay_block_file(block_name='DemoReplayBlock', recording_file_path=fr'{self.file_path}\demo_cfg\Logs\demo_log.blf')
        wait(1)
        self.canoe_inst.control_replay_block(block_name='DemoReplayBlock', start_stop=True)
        wait(2)
        self.canoe_inst.control_replay_block(block_name='DemoReplayBlock', start_stop=False)
        wait(1)
        assert self.canoe_inst.stop_measurement()

    def test_capl_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=True, auto_save=True, prompt_user=False, auto_stop=True)
        self.canoe_inst.compile_all_capl_nodes()
        assert self.canoe_inst.start_measurement()
        wait(1)
        assert self.canoe_inst.call_capl_function('addition_function', 100, 200)
        assert self.canoe_inst.call_capl_function('hello_world')
        assert self.canoe_inst.stop_measurement()

    def test_test_setup_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_test_setup)
        self.canoe_inst.ui_activate_desktop('TestSetup')
        assert self.canoe_inst.start_measurement()
        wait(1)
        self.canoe_inst.execute_all_test_environments()
        test_environments = self.canoe_inst.get_test_environments()
        for te_name, _ in test_environments.items():
            self.canoe_inst.execute_all_test_modules_in_test_env(te_name)
        self.canoe_inst.execute_test_module('demo_test_node_001')
        self.canoe_inst.execute_test_module('demo_test_node_002')
        wait(1)
        assert self.canoe_inst.stop_measurement()

    def test_env_var_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev)
        assert self.canoe_inst.start_measurement()
        wait(1)
        self.canoe_inst.set_environment_variable_value('int_var', 123.12)
        self.canoe_inst.get_environment_variable_value('int_var')
        self.canoe_inst.set_environment_variable_value('float_var', 111.123)
        self.canoe_inst.get_environment_variable_value('float_var')
        self.canoe_inst.set_environment_variable_value('string_var', 'this is string variable')
        self.canoe_inst.get_environment_variable_value('string_var')
        self.canoe_inst.set_environment_variable_value('data_var', (1, 2, 3, 4, 5, 6, 7))
        self.canoe_inst.get_environment_variable_value('data_var')
        wait(1)
        assert self.canoe_inst.stop_measurement()

    def test_conf_gen_setup(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_gen_db_setup, visible=True, auto_save=True, prompt_user=False, auto_stop=True)
        assert self.canoe_inst.start_measurement()
        wait(1)
        self.canoe_inst.add_database(fr"{self.file_path}\demo_cfg\DBs\sample_databases\XCP.dbc", 'CAN1', 1)
        self.canoe_inst.remove_database(fr"{self.file_path}\demo_cfg\DBs\sample_databases\XCP.dbc", 1)
        assert self.canoe_inst.stop_measurement()
        assert self.canoe_inst.add_database(fr"{self.file_path}\demo_cfg\DBs\sample_databases\XCP.dbc", 'CAN1', 1)
        assert self.canoe_inst.remove_database(fr"{self.file_path}\demo_cfg\DBs\sample_databases\XCP.dbc", 1)
        assert self.canoe_inst.save_configuration()

    def test_online_setup_logging(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_online_setup)
        assert self.canoe_inst.start_measurement()
        wait(1)
        self.canoe_inst.start_stop_online_measurement_setup_logging_block(
            'd:\\git_repos\\chaitu-ycr\\py_canoe\\tests\\demo_cfg\\Logs\\demo_online_setup_log.blf',
            start=False,
        )
        wait(2)
        self.canoe_inst.start_stop_online_measurement_setup_logging_block(
            'd:\\git_repos\\chaitu-ycr\\py_canoe\\tests\\demo_cfg\\Logs\\demo_online_setup_log.blf',
            start=True,
        )
        wait(2)
        self.canoe_inst.start_stop_online_measurement_setup_logging_block(
            'd:\\git_repos\\chaitu-ycr\\py_canoe\\tests\\demo_cfg\\Logs\\demo_online_setup_log.blf',
            start=False,
        )
        assert self.canoe_inst.stop_measurement()


class TestPyCanoeFlask:
    @classmethod
    def setup_class(cls):
        cls.file_path = os.path.dirname(os.path.abspath(__file__))
        cls.root_path = cls.file_path
        cls.canoe_inst = CANoe(py_canoe_log_dir=os.path.join(cls.root_path, ".py_canoe_log"), user_capl_functions=('addition_function', 'hello_world'))
        demo_cfg_dir = os.path.join(cls.file_path, "demo_cfg")
        cls.canoe_cfg_one_ch = os.path.join(demo_cfg_dir, "demo_can_one_ch.cfg")
        cls.canoe_cfg_two_ch = os.path.join(demo_cfg_dir, "demo_can_two_ch.cfg")
        cls.canoe_cfg_gen_db_setup = os.path.join(demo_cfg_dir, "demo_conf_gen_db_setup.cfg")
        cls.canoe_cfg_dev = os.path.join(demo_cfg_dir, "demo_dev.cfg")
        cls.canoe_cfg_diag = os.path.join(demo_cfg_dir, "demo_diag.cfg")
        cls.canoe_cfg_eth_one_ch = os.path.join(demo_cfg_dir, "demo_eth_one_ch.cfg")
        cls.canoe_cfg_offline = os.path.join(demo_cfg_dir, "demo_offline.cfg")
        cls.canoe_cfg_test_setup = os.path.join(demo_cfg_dir, "demo_test_setup.cfg")
        cls.canoe_cfg_demo = os.path.join(demo_cfg_dir, "demo.cfg")

    def test_meas_start_stop_restart_methods(self):
        self.canoe_inst.open(canoe_cfg=self.canoe_cfg_dev, visible=True, auto_save=False, prompt_user=False)
        assert self.canoe_inst.start_measurement()
        assert self.canoe_inst.stop_measurement()
        assert self.canoe_inst.start_measurement()
        assert self.canoe_inst.reset_measurement()
        assert self.canoe_inst.get_measurement_running_status()
        assert self.canoe_inst.stop_ex_measurement()
        assert not self.canoe_inst.get_measurement_running_status()
        self.canoe_inst.quit()