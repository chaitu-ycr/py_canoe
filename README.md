# About [py_canoe](https://github.com/chaitu-ycr/py_canoe)

Python 🐍 Package for controlling Vector CANoe 🛶 Tool

## Acknowledgements

I want to thank plants 🎋 for providing me oxygen each day.
Also, I want to thank the sun 🌄 for providing more than half of their nourishment free of charge.

## Prerequisites

- [X] Python(>=3.9)
- [X] Vector CANoe software(>=v11)
- [X] Windows PC(recomended win 10 os)
- [X] visual studio code

## python environment setup

create python virtual environment

```bat
python -m venv .venv
```

activate virtual environment

```bat
.venv\Scripts\activate
```

upgrade pip

```bat
python -m pip install pip --upgrade
```

install [py_canoe](https://pypi.org/project/py_canoe/) module

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

For all available CANoe methods check user documentation.

### User Documentation [click here](https://chaitu-ycr.github.io/py_canoe/)
