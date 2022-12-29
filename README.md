# py_canoe
Python Library for controlling Vector CANoe tool

## Installation
Always Create Virtual environment
```bat
python -m venv venv
```
Activate virtual environment and upgrade pip
```bat
venv\Scripts\activate
pip install pip --upgrade
```
Install py_canoe module
```bat
pip install py_canoe
```
## Usage
### Import CANoe module
```python
# Import CANoe module
from py_canoe import CANoe

# create CANoe object
canoe_inst = CANoe()
```
### Some more commonly used methods
```python
# open CANoe configuration. Replace canoe_cfg with yours.
canoe_inst.open(canoe_cfg=r'C:\Users\Public\Documents\Vector\CANoe\Sample Configurations 11.0.81\.\CAN\Diagnostics\UDSBasic\UDSBasic.cfg')

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
resp = canoe_inst.send_diag_request('CAN', 'Door', '10 01')
print(resp)

# Stop CANoe Measurement
canoe_inst.stop_measurement()

# Quit / Close CANoe configuration
canoe_inst.quit()
```

For all available CANoe methods check user documentation.

### User Documentation [click here](https://chaitu-ycr.github.io/py_canoe/)
