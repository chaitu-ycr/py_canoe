import os
import logging
from time import sleep as wait
from py_canoe import CANoe

file_path = os.path.dirname(os.path.abspath(__file__)).replace('/', '\\')
root_path = file_path
canoe_inst = CANoe(py_canoe_log_dir=fr'{root_path}\.py_canoe_log', user_capl_functions=('addition_function', 'hello_world'))
logger_inst = logging.getLogger('CANOE_LOG')


def test_dummy_test_001():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_conf_gen_db_setup.cfg', visible=True, auto_save=False, prompt_user=False)
    channels = canoe_inst.application.configuration.general_setup.get_channels(1)
    canoe_inst.application.configuration.general_setup.set_channels(1, 3)
    channels = canoe_inst.application.configuration.general_setup.get_channels(1)

def test_dummy_test_002():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_conf_gen_db_setup.cfg', visible=True, auto_save=False, prompt_user=False)
    databases = canoe_inst.application.configuration.general_setup.database_setup.fetch_all_databases()
    databases['Comfort'].fullname = 'd:\\git_repos\\py_canoe_dev\\tests\\demo_cfg\\DBs\\sample_databases\\PowerTrain.dbc'
    databases = canoe_inst.application.configuration.general_setup.database_setup.fetch_all_databases()
    channels = canoe_inst.application.configuration.general_setup.get_channels(1)