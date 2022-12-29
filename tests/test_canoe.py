import os
import sys

file_path = os.path.dirname(os.path.abspath(__file__)).replace('/', '\\')
root_path = '\\'.join(file_path.split('\\')[:-1])
sys.path.extend([file_path, fr'{root_path}\src'])

from py_canoe import CANoe
canoe_inst = CANoe()


def test_check_canoe_measurement():
    canoe_inst.open(fr'{file_path}\demo_cfg\demo.cfg')
    canoe_inst.get_canoe_version_info()
    canoe_inst.start_measurement()
    assert canoe_inst.get_measurement_running_status()
    canoe_inst.stop_measurement()
    assert not canoe_inst.get_measurement_running_status()
    canoe_inst.quit()
