import win32com.client
from typing import Union

from py_canoe.utils.common import DoEventsUntil
from py_canoe.utils.common import logger


class VariablesFile:
    def __init__(self, variables_file):
        self.com_object = variables_file

    @property
    def full_name(self) -> str:
        return self.com_object.FullName

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def path(self) -> str:
        return self.com_object.Path


class VariablesFiles:
    def __init__(self, variables_files):
        self.com_object = variables_files

    @property
    def count(self):
        return self.com_object.Count

    def item(self, index: int) -> VariablesFile:
        return VariablesFile(self.com_object.Item(index))

    def add(self, name: str):
        return self.com_object.Add(name)

    def remove(self, variable_file):
        return self.com_object.Remove(variable_file)


class Encoding:
    def __init__(self, encoding):
        self.com_object = encoding

    @property
    def factor(self) -> float:
        return self.com_object.Factor

    @property
    def lower_bound(self) -> int:
        return self.com_object.LowerBound

    @property
    def offset(self) -> float:
        return self.com_object.Offset

    @property
    def text(self) -> str:
        return self.com_object.Text

    @property
    def unit(self) -> str:
        return self.com_object.Unit

    @property
    def upper_bound(self) -> int:
        return self.com_object.UpperBound


class Encodings:
    def __init__(self, encodings):
        self.com_object = encodings

    @property
    def count(self):
        return self.com_object.Count

    def item(self, index: int) -> Encoding:
        return Encoding(self.com_object.Item(index))


class VariableEvents:
    def __init__(self) -> None:
        self.VARIABLE_INFO = {}
        self.VARIABLE_UPDATED = False

    def OnChange(self, value):
        self.VARIABLE_INFO['value'] = value
        self.VARIABLE_UPDATED = True

    def OnChangeAndTime(self, value, timeHigh, time):
        self.VARIABLE_INFO['value'] = value
        self.VARIABLE_INFO['timeHigh'] = timeHigh
        self.VARIABLE_INFO['time'] = time
        self.VARIABLE_UPDATED = True


class Variable:
    def __init__(self, variable, enable_events: bool = True):
        self.com_object = win32com.client.Dispatch(variable)
        self.enable_events = enable_events
        if self.enable_events:
            self.variable_events: VariableEvents = win32com.client.WithEvents(self.com_object, VariableEvents)

    @property
    def analysis_only(self) -> bool:
        return self.com_object.AnalysisOnly

    @analysis_only.setter
    def analysis_only(self, value: bool = False):
        self.com_object.AnalysisOnly = value

    @property
    def bit_count(self) -> int:
        return self.com_object.BitCount

    @property
    def comment(self) -> str:
        return self.com_object.Comment

    @property
    def element_count(self) -> int:
        return self.com_object.ElementCount

    @property
    def encodings(self) -> Encodings:
        return Encodings(self.com_object.Encodings)

    @property
    def full_name(self) -> str:
        return self.com_object.FullName

    @property
    def init_value(self) -> Union[int, float]:
        return self.com_object.InitValue

    @property
    def is_array(self) -> bool:
        return self.com_object.IsArray

    @property
    def is_member(self) -> bool:
        return self.com_object.IsMember

    @property
    def is_signed(self) -> bool:
        return self.com_object.IsSigned

    @property
    def is_struct(self) -> bool:
        return self.com_object.IsStruct

    @property
    def max_value(self) -> Union[int, float]:
        return self.com_object.MaxValue

    @property
    def member_name(self) -> str:
        return self.com_object.MemberName

    @property
    def members(self) -> 'Variables':
        return Variables(self.com_object.Members)

    @property
    def min_value(self) -> Union[int, float]:
        return self.com_object.MinValue

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def notification_type(self) -> int:
        return self.com_object.NotificationType

    @property
    def physical_init_value(self) -> Union[int, float]:
        return self.com_object.PhysicalInitValue

    @property
    def physical_value(self) -> Union[int, float]:
        return self.com_object.PhysicalValue

    @property
    def read_only(self) -> bool:
        return self.com_object.ReadOnly

    @property
    def type(self) -> int:
        return self.com_object.Type

    @property
    def unit(self) -> str:
        return self.com_object.Unit

    def get_value(self):
        return self.com_object.Value

    def set_value(self, value, timeout: Union[int, float]):
        status: bool = False
        self.com_object.Value = value
        if self.variable_events:
            self.variable_events.VARIABLE_UPDATED = False
            status = DoEventsUntil(lambda: self.variable_events.VARIABLE_UPDATED, timeout, "Variable Update")
            if status:
                logger.info(f"📢 Variable '{self.full_name}' updated successfully to: {value}")
        return status

    def begin_struct_update(self):
        self.com_object.BeginStructUpdate()

    def end_struct_update(self):
        self.com_object.EndStructUpdate()

    def get_member_phys_value(self, member_name: str) -> Union[int, float]:
        return self.com_object.GetMemberPhysValue(member_name)

    def get_member_value(self, member_name: str) -> Union[int, float]:
        return self.com_object.GetMemberValue(member_name)

    def get_symbolic_value_name(self, value: Union[int, float]) -> str:
        return self.com_object.GetSymbolicValueName(value)

    def set_member_phys_value(self, member_name: str, value: Union[int, float]):
        self.com_object.SetMemberPhysValue(member_name, value)

    def set_member_value(self, member_name: str, value: Union[int, float]):
        self.com_object.SetMemberValue(member_name, value)

    def set_symbolic_value_name(self, value: Union[int, float], name: str):
        self.com_object.SetSymbolicValueName(value, name)


class Variables:
    def __init__(self, variables):
        self.com_object = variables

    @property
    def count(self):
        return self.com_object.Count

    def item(self, index: int) -> Variable:
        return Variable(self.com_object.Item(index))

    def add(self, name: str):
        return self.com_object.Add(name)

    def add_ex(self, name: str, initial_value: Union[int, float], min_value: Union[int, float], max_value: Union[int, float]):
        return self.com_object.AddEx(name, initial_value, min_value, max_value)

    def add_writeable(self, name: str, initial_value: Union[int, float]):
        return self.com_object.AddWriteable(name, initial_value)

    def add_writable_ex(self, name: str, initial_value: Union[int, float], min_value: Union[int, float] = None, max_value: Union[int, float] = None):
        return self.com_object.AddWritableEx(name, initial_value, min_value, max_value)

    def remove(self, variable):
        return self.com_object.Remove(variable)


class Namespace:
    def __init__(self, namespace):
        self.com_object = namespace

    @property
    def comment(self):
        return self.com_object.Comment

    @property
    def name(self):
        return self.com_object.Name

    def variables(self) -> Variables:
        return Variables(self.com_object.Variables)

    def namespaces(self) -> 'Namespaces':
        return Namespaces(self.com_object.Namespaces)


class Namespaces:
    def __init__(self, namespaces):
        self.com_object = namespaces

    @property
    def count(self):
        return self.com_object.Count

    def item(self, index: int) -> Namespace:
        return Namespace(self.com_object.Item(index))

    def add(self, name: str):
        return self.com_object.Add(name)

    def remove(self, variable):
        return self.com_object.Remove(variable)


class System:
    """
    The System object represents the system of the CANoe application.
    The System object offers access to the namespaces for data exchange with external applications.
    """
    def __init__(self, app):
        self.com_object = app.com_object.System

    @property
    def namespaces(self) -> Namespaces:
        return Namespaces(self.com_object.Namespaces)

    @property
    def variables_files(self) -> VariablesFiles:
        return VariablesFiles(self.com_object.VariablesFiles)


    def add_system_variable(self, sys_var_name: str, value: Union[int, float, str], read_only: bool = False) -> Union[object, None]:
        new_var_com_obj = None
        try:
            parts = sys_var_name.split('::')
            if len(parts) < 2:
                logger.error(f"❌ Invalid system variable name '{sys_var_name}'. Must be in 'namespace::variable' format.")
                return None
            namespace = '::'.join(parts[:-1])
            variable_name = parts[-1]
            try:
                namespace_obj = self.com_object.Namespaces(namespace)
            except Exception:
                logger.info(f"namespace '{namespace}' not present. Creating namespace...")
                namespaces_obj = self.com_object.Namespaces
                namespace_obj = namespaces_obj.Add(namespace)
                logger.info(f"Created new namespace: {namespace}")
            variables_obj = namespace_obj.Variables
            if read_only:
                new_var_com_obj = variables_obj.Add(variable_name, value)
            else:
                new_var_com_obj = variables_obj.AddWriteable(variable_name, value)
            logger.info(f"System Variable '{sys_var_name}' defined successfully with value: {value}")
            return new_var_com_obj
        except Exception as e:
            logger.error(f"❌ Error defining System Variable '{sys_var_name}': {e}")
            return None

    def remove_system_variable(self, sys_var_name: str) -> bool:
        try:
            parts = sys_var_name.split('::')
            if len(parts) < 2:
                logger.error(f"❌ Invalid system variable name '{sys_var_name}'. Must be in 'namespace::variable' format.")
                return None
            namespace = '::'.join(parts[:-1])
            variable_name = parts[-1]
            namespace_obj = self.com_object.Namespaces(namespace)
            variables_obj = namespace_obj.Variables
            for i in range(1, variables_obj.Count + 1):
                variable_obj = variables_obj.Item(i)
                if variable_obj.Name == variable_name:
                    variables_obj.Remove(i)
                    logger.info(f"System Variable '{sys_var_name}' removed successfully.")
                    return True
            logger.info(f"System Variable '{sys_var_name}' not found.")
            return False
        except Exception as e:
            logger.error(f"❌ Error removing System Variable '{sys_var_name}': {e}")
            return False

    def get_system_variable_value(self, sys_var_name: str, return_symbolic_name=False) -> Union[int, float, str, None]:
        try:
            parts = sys_var_name.split('::')
            if len(parts) < 2:
                logger.error(f"❌ Invalid system variable name '{sys_var_name}'. Must be in 'namespace::variable' format.")
                return None
            namespace = '::'.join(parts[:-1])
            variable_name = parts[-1]
            namespace_obj = self.com_object.Namespaces(namespace)
            variable_obj = Variable(namespace_obj.Variables(variable_name))
            value = variable_obj.get_value()
            if return_symbolic_name:
                symbolic_value = variable_obj.get_symbolic_value_name(value)
                logger.info(f"System Variable '{sys_var_name}' symbolic value: {symbolic_value}")
                return symbolic_value
            logger.info(f"System Variable '{sys_var_name}' value: {value}")
            return value
        except Exception as e:
            logger.error(f"❌ Error retrieving System Variable '{sys_var_name}': {e}")
            return None

    def set_system_variable_value(self, sys_var_name: str, value: Union[int, float, str], timeout: Union[int, float] = 1) -> bool:
        try:
            parts = sys_var_name.split('::')
            if len(parts) < 2:
                logger.error(f"❌ Invalid system variable name '{sys_var_name}'. Must be in 'namespace::variable' format.")
                return False
            namespace = '::'.join(parts[:-1])
            variable_name = parts[-1]
            namespace_obj = self.com_object.Namespaces(namespace)
            variable_obj = Variable(namespace_obj.Variables(variable_name))
            var_type = type(variable_obj.get_value())
            try:
                converted_value = var_type(value)
            except Exception:
                logger.error(f"❌ Could not convert value '{value}' to type {var_type.__name__} for '{sys_var_name}'")
                return False
            status = variable_obj.set_value(converted_value, timeout)
            return status
        except Exception as e:
            logger.error(f"❌ Error setting System Variable '{sys_var_name}': {e}")
            return False

    def set_system_variable_array_values(self, sys_var_name: str, value: tuple, index: int = 0, timeout: Union[int, float] = 1) -> bool:
        try:
            parts = sys_var_name.split('::')
            if len(parts) < 2:
                logger.error(f"❌ Invalid system variable name '{sys_var_name}'. Must be in 'namespace::variable' format.")
                return False
            namespace = '::'.join(parts[:-1])
            variable_name = parts[-1]
            namespace_obj = self.com_object.Namespaces(namespace)
            variable_obj = Variable(namespace_obj.Variables(variable_name))
            arr = list(variable_obj.get_value())
            if index < 0 or index + len(value) > len(arr):
                logger.error(f"❌ Not enough space in System Variable Array '{sys_var_name}' to set values.")
                return False
            value_type = type(arr[0]) if arr else type(value[0])
            arr[index:index + len(value)] = [value_type(v) for v in value]
            status = variable_obj.set_value(tuple(arr), timeout)
            return status
        except Exception as e:
            logger.error(f"❌ Error setting System Variable Array '{sys_var_name}': {e}")
            return False
