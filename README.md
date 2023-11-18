# About [py_canoe](https://github.com/chaitu-ycr/py_canoe)

Python 🐍 Package for controlling Vector CANoe 🛶 Tool

fork [this repo](https://github.com/chaitu-ycr/py_canoe/fork) and create pull request to contribute back to this project.

for ideas/discussions please create new discussion [here](https://github.com/chaitu-ycr/py_canoe/discussions)

create issue or request feature [here](https://github.com/chaitu-ycr/py_canoe/issues/new/choose)

## GitHub Releases 👉 [link](https://github.com/chaitu-ycr/py_canoe/releases)

## PyPi Package 👉 [link](https://pypi.org/project/py_canoe/)

## Prerequisites [link](https://chaitu-ycr.github.io/py_canoe/002_prerequisites/)

## Python environment setup instructions [link](https://chaitu-ycr.github.io/py_canoe/003_environment_setup/)

## Install [py_canoe](https://pypi.org/project/py_canoe/) package

```bat
pip install py_canoe --upgrade
```

## Documentation Links

### example use cases 👉 [link](https://chaitu-ycr.github.io/py_canoe/004_usage/)

### package reference doc 👉 [link](https://chaitu-ycr.github.io/py_canoe/999_reference/)

## Sample Example

```python
# Import CANoe module
from py_canoe import CANoe

# create CANoe object
canoe_inst = CANoe()

# open CANoe configuration. Replace canoe_cfg with yours.
canoe_inst.open(canoe_cfg=r'tests\demo_cfg\demo.cfg')

# print installed CANoe application version
canoe_inst.get_canoe_version_info()

# Start CANoe measurement
canoe_inst.start_measurement()

# get signal value. Replace arguments with your message and signal data.
sig_val = canoe_inst.get_signal_value('CAN', 1, 'LightState', 'FlashLight')
print(sig_val)

# Stop CANoe Measurement
canoe_inst.stop_measurement()

# Quit / Close CANoe configuration
canoe_inst.quit()
```
