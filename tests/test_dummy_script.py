import os
import logging
from time import sleep as wait
from py_canoe import CANoe

file_path = os.path.dirname(os.path.abspath(__file__)).replace('/', '\\')
root_path = file_path
canoe_inst = CANoe(py_canoe_log_dir=fr'{root_path}\.py_canoe_log')
logger_inst = logging.getLogger('CANOE_LOG')


def test_dummy_test_001():
    canoe_inst.open(canoe_cfg=fr'{file_path}\demo_cfg\demo_conf_gen_db_setup.cfg', visible=True, auto_save=False, prompt_user=False)

def test_dummy_test_002():
    canoe_inst.open(canoe_cfg=fr'C:\Users\Public\Documents\Vector\CANoe\Sample Configurations 11.0.96\Ethernet\EthernetSystem\EthernetSystem.cfg',
                    visible=True, auto_save=False, prompt_user=False)
    canoe_inst.start_measurement()
    wait(2)
    canoe_inst.get_signal_value(bus='Ethernet', channel=1, message='VehicleSpeed', signal='VehicleSpeed')
    canoe_inst.stop_measurement()