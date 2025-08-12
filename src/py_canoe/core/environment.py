# TODO: complete implementation of the Environment class
import win32com.client
from typing import Union

from py_canoe.utils.common import DoEventsUntil


class EnvironmentVariableEvents:
    VARIABLE_CHANGED = False
    VARIABLE_EVENTS_INFO = {}

    @staticmethod
    def OnChange(value):
        EnvironmentVariableEvents.VARIABLE_CHANGED = True
        EnvironmentVariableEvents.VARIABLE_EVENTS_INFO['value'] = value
    
    @staticmethod
    def OnChangeAndTime(value, timeHigh, time):
        EnvironmentVariableEvents.VARIABLE_CHANGED = True
        EnvironmentVariableEvents.VARIABLE_EVENTS_INFO['value'] = value
        EnvironmentVariableEvents.VARIABLE_EVENTS_INFO['timeHigh'] = timeHigh
        EnvironmentVariableEvents.VARIABLE_EVENTS_INFO['time'] = time
    
    @staticmethod
    def OnChangeAndTimeU(value, timeHigh, time):
        EnvironmentVariableEvents.VARIABLE_CHANGED = True
        EnvironmentVariableEvents.VARIABLE_EVENTS_INFO['value'] = value
        EnvironmentVariableEvents.VARIABLE_EVENTS_INFO['timeHigh'] = timeHigh
        EnvironmentVariableEvents.VARIABLE_EVENTS_INFO['time'] = time


class EnvironmentVariable:
    def __init__(self, environment_variable, enable_events: bool = True):
        self.com_object = win32com.client.Dispatch(environment_variable)
        self.value_change_timeout = 1
        EnvironmentVariableEvents.VARIABLE_CHANGED = False
        self.wait_for_change = lambda: DoEventsUntil(lambda: EnvironmentVariableEvents.VARIABLE_CHANGED, self.value_change_timeout, "Environment Variable Change")
        if enable_events:
            win32com.client.WithEvents(self.com_object, EnvironmentVariableEvents)

    @property
    def handle(self) -> int:
        return self.com_object.Handle
    
    @handle.setter
    def handle(self, value: int):
        self.com_object.Handle = value

    @property
    def notification_type(self) -> int:
        return self.com_object.NotificationType
    
    @property
    def type(self) -> int:
        return self.com_object.Type
    
    @property
    def value(self) -> Union[str, int, float]:
        return self.com_object.Value
    
    @value.setter
    def value(self, value: Union[str, int, float]):
        self.com_object.Value = value
        self.wait_for_change()


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
