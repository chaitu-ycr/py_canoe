# import external modules here
import logging
import pythoncom
import win32com.client
from time import sleep as wait


class Environment:
    """The Environment class represents the environment variables.
    The Environment class is only available in CANoe
    """
    def __init__(self, app_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.Environment)
        except Exception as e:
            self.__log.error(f'😡 Error initializing Environment object: {str(e)}')

    def create_group(self):
        return EnvironmentGroup(self.com_obj.CreateGroup())

    def create_info(self):
        return EnvironmentInfo(self.com_obj.CreateInfo())

    def get_variable(self, name: str):
        return EnvironmentVariable(self.com_obj.GetVariable(name))

    def get_variables(self, list_of_variable_names: tuple):
        return self.com_obj.GetVariables(list_of_variable_names)

    def set_variables(self, list_of_variables_with_name_value: tuple):
        self.com_obj.SetVariables(list_of_variables_with_name_value)


class EnvironmentGroup:
    """The EnvironmentGroup class represents a group of environment variables.
    With the help of environment variable groups you can set or query multiple environment variables simultaneously with just one call.
    """
    def __init__(self, env_group_com_obj):
        self.com_obj = env_group_com_obj

    @property
    def array(self):
        return EnvironmentArray(self.com_obj.Array)

    def add(self, variable):
        self.com_obj.Add(variable)

    def get_values(self):
        return self.com_obj.GetValues()

    def remove(self, variable):
        self.com_obj.Variable(variable)

    def set_values(self, values):
        self.com_obj.SetValues(values)


class EnvironmentArray:
    """The EnvironmentArray class represents an array of environment variables."""
    def __init__(self, env_array_com_obj):
        self.com_obj = env_array_com_obj

    @property
    def count(self) -> int:
        return self.com_obj.Count


def DoEvents():
    pythoncom.PumpWaitingMessages()
    wait(.1)

def DoEventsUntil(condition):
    while not condition():
        DoEvents()


class EnvironmentVariableEvents:
    def __init__(self):
        self.var_event_occurred = False

    def OnChange(self, value):
        self.var_event_occurred = True


class EnvironmentVariable:
    """The EnvironmentVariable class represents an environment variable."""
    def __init__(self, env_var_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.DispatchWithEvents(env_var_com_obj, EnvironmentVariableEvents)
            self.wait_for_tm_to_start = lambda: DoEventsUntil(lambda: self.com_obj.var_event_occurred)
        except Exception as e:
            self.__log.error(f'😡 Error initializing EnvironmentVariable object: {str(e)}')

    @property
    def handle(self):
        return self.com_obj.Handle

    @handle.setter
    def handle(self, value):
        self.com_obj.Handle = value

    @property
    def notification_type(self):
        return self.com_obj.NotificationType

    @notification_type.setter
    def notification_type(self, value: int):
        self.com_obj.NotificationType = value

    @property
    def type(self):
        return self.com_obj.Type

    @property
    def value(self):
        return self.com_obj.Value

    @value.setter
    def value(self, value):
        self.com_obj.Value = value
        wait(.1)


class EnvironmentInfo:
    """The EnvironmentInfo class represents information related to name, type and associated database of environment variables.
    With the environment information you can determine how the access to environment variables is configured (writable, readable, both).
    """
    def __init__(self, env_info_com_obj):
        self.com_obj = env_info_com_obj

    @property
    def read(self):
        return self.com_obj.Read

    @property
    def write(self):
        return self.com_obj.Write

    def get_info(self):
        return self.com_obj.GetInfo()
