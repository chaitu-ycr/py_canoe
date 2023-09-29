# Usage

## Import CANoe module

```python
# Import CANoe module
from py_canoe import CANoe

# create CANoe object
canoe_inst = CANoe()
```

## Example use cases

```python
# open CANoe configuration. Replace canoe_cfg with yours.
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo.cfg')

# print installed CANoe application version
canoe_inst.get_canoe_version_info()

# Start CANoe measurement
canoe_inst.start_measurement()

# get signal value. Replace arguments with your message and signal data.
sig_val = canoe_inst.get_signal_value('CAN', 1, 'LightState', 'FlashLight')
print(sig_val)

# set signal value. Replace arguments with your message and signal data.
canoe_inst.set_signal_value('CAN', 1, 'LightState', 'FlashLight', 1)

# send diagnostic request. Replace arguments with your diagnostics data.
resp = canoe_inst.send_diag_request('Door', '10 01')
print(resp)

# Stop CANoe Measurement
canoe_inst.stop_measurement()

# Quit / Close CANoe configuration
canoe_inst.quit()
```
