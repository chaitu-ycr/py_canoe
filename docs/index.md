# py_canoe

Python 🐍 Package for controlling Vector CANoe 🛶 Tool

## Acknowledgements

I want to thank plants 🎋 for providing me oxygen each day.
Also, I want to thank the sun 🌄 for providing more than half of their nourishment free of charge.

## Prerequisites

- [X] Python(>=3.6)
- [X] Vector CANoe software(>=v11)
- [X] Windows PC(recomended win 10 os)

## Installation

### Create Virtual environment

```bat
python -m venv venv
```

Activate virtual environment and upgrade pip

```bat
venv\Scripts\activate
python -m pip install pip --upgrade
```

### Install py_canoe module

```bat
pip install py_canoe --upgrade
```

## Usage

### Import CANoe module

```python
# Import CANoe module
from py_canoe import CANoe

# create CANoe object
canoe_inst = CANoe()
```

### Example use cases

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

## CANoe class reference list

### ::: src.py_canoe.CANoe
