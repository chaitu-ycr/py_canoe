# TODO: complete implementation of the Environment class
import win32com.client
from typing import Union

from py_canoe.utils.common import DoEventsUntil, logger

ENV_VAR_CHANGE_TIMEOUT = 1


class EnvironmentVariable:
    def __init__(self, environment_variable):
        self.com_object = environment_variable
        self.VALUE_TABLE_NOTIFICATION_TYPE = {
            0: "cValue",
            1: "cValueAndTime",
            2: "cValueAndTimeU"
        }
        self.VALUE_TABLE_TYPE = {
            0: "INTEGER",
            1: "FLOAT",
            2: "STRING",
            3: "DATA"
        }

    @property
    def handle(self) -> int:
        return self.com_object.Handle

    @handle.setter
    def handle(self, value: int):
        self.com_object.Handle = value

    @property
    def notification_type(self) -> int:
        return self.com_object.NotificationType

    @notification_type.setter
    def notification_type(self, value: int):
        self.com_object.NotificationType = value

    @property
    def type(self) -> int:
        return self.com_object.Type

    @property
    def value(self) -> Union[str, int, float]:
        return self.com_object.Value

    @value.setter
    def value(self, value: Union[str, int, float]):
        self.com_object.Value = value
        DoEventsUntil(lambda: self._check_value_updated(value), ENV_VAR_CHANGE_TIMEOUT, "Environment Variable Change")

    def _check_value_updated(self, value) -> bool:
        set_value = value
        get_value = self.value if self.type != 3 else tuple(self.value)
        return get_value == set_value


class EnvironmentArray:
    def __init__(self, environment_array):
        self.com_object = environment_array

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> EnvironmentVariable:
        return EnvironmentVariable(self.com_object.Item(index))


class EnvironmentGroup:
    def __init__(self, environment_group):
        self.com_object = environment_group

    @property
    def array(self) -> EnvironmentArray:
        return EnvironmentArray(self.com_object.Array)

    def add(self, variable):
        self.com_object.Add(variable)

    def get_values(self):
        return self.com_object.GetValues()

    def remove(self, variable):
        self.com_object.Remove(variable)

    def set_values(self, values: list):
        self.com_object.SetValues(values)


class EnvironmentInfo:
    def __init__(self, environment_info):
        self.com_object = environment_info

    @property
    def read(self) -> bool:
        return self.com_object.Read

    @property
    def write(self) -> bool:
        return self.com_object.Write

    def get_info(self) -> list:
        return self.com_object.GetInfo()


class Environment:
    """
    The Environment object represents the environment variables.
    """
    def __init__(self, app):
        self.com_object = app.com_object.Environment

    def create_group(self):
        return EnvironmentGroup(self.com_object.CreateGroup())

    def create_info(self) -> EnvironmentInfo:
        return EnvironmentInfo(self.com_object.CreateInfo())

    def get_variable(self, name: str) -> EnvironmentVariable:
        return EnvironmentVariable(self.com_object.GetVariable(name))

    def get_variables(self, vars: list[list[Union[str, int, float]]]) -> list:
        return self.com_object.GetVariables(vars)

    def set_variables(self, vars: dict):
        self.com_object.SetVariables(vars)

    def get_environment_variable_value(self, env_var_name: str) -> Union[int, float, str, tuple, None]:
        var_value = None
        try:
            variable = self.get_variable(env_var_name)
            var_value = variable.value if variable.type != 3 else tuple(variable.value)
            logger.info(f'🔢 environment variable({env_var_name}) value = {var_value}')
        except Exception as e:
            logger.error(f"❌ Failed to get environment variable '{env_var_name}': {e}")
        finally:
            return var_value

    def set_environment_variable_value(self, env_var_name: str, value: Union[int, float, str, tuple]) -> bool:
        try:
            variable = self.get_variable(env_var_name)
            if variable.type == 0:
                converted_value = int(value)
            elif variable.type == 1:
                converted_value = float(value)
            elif variable.type == 2:
                converted_value = str(value)
            else:
                converted_value = tuple(value)
            variable.value = converted_value
            logger.info(f'🔢 environment variable({env_var_name}) set to 👉 {converted_value}')
            return True
        except Exception as e:
            logger.error(f"❌ Failed to set environment variable '{env_var_name}': {e}")
            return False
