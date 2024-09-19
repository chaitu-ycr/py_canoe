# import external modules here
import os
import logging
import win32com.client

# import internal modules here


class System:
    """The System object represents the system of the CANoe application.
    The System object offers access to the namespaces for data exchange with external applications.
    """
    def __init__(self, app_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.System)
            self.namespaces_com_obj = win32com.client.Dispatch(self.com_obj.Namespaces)
            self.variables_files_com_obj = win32com.client.Dispatch(self.com_obj.VariablesFiles)
            self.namespaces_dict = {}
            self.variables_files_dict = {}
            self.variables_dict = {}
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe System: {str(e)}')

    @property
    def namespaces_count(self) -> int:
        return self.namespaces_com_obj.Count

    def fetch_namespaces(self) -> dict:
        if self.namespaces_count > 0:
            for index in range(1, self.namespaces_count + 1):
                namespace_com_obj = win32com.client.Dispatch(self.namespaces_com_obj.Item(index))
                namespace_name = namespace_com_obj.Name
                self.namespaces_dict[namespace_name] = namespace_com_obj
                if 'Namespaces' in dir(namespace_com_obj):
                    self.fetch_namespace_namespaces(namespace_com_obj, namespace_name)
                if 'Variables' in dir(namespace_com_obj):
                    self.fetch_namespace_variables(namespace_com_obj)
        return self.namespaces_dict

    def add_namespace(self, name: str):
        self.fetch_namespaces()
        if name not in self.namespaces_dict.keys():
            namespace_com_obj = self.namespaces_com_obj.Add(name)
            self.namespaces_dict[name] = namespace_com_obj
            self.__log.debug(f'Added the new namespace ({name}).')
            return namespace_com_obj
        else:
            self.__log.warning(f'The given namespace ({name}) already exists.')
            return None

    def remove_namespace(self, name: str) -> None:
        self.fetch_namespaces()
        if name in self.namespaces_list:
            self.namespaces_com_obj.Remove(name)
            self.fetch_namespaces()
            self.__log.debug(f'Removed the namespace ({name}) from the collection.')
        else:
            self.__log.warning(f'The given namespace ({name}) does not exist.')

    @property
    def varaibles_files_count(self) -> int:
        return self.variables_files_com_obj.Count

    def fetch_variables_files(self):
        if self.varaibles_files_count > 0:
            for index in range(1, self.varaibles_files_count + 1):
                variable_file_com_obj = self.variables_files_com_obj.Item(index)
                self.variables_files_dict[variable_file_com_obj.Name] = {'full_name': variable_file_com_obj.FullName,
                                                                         'path': variable_file_com_obj.Path,
                                                                         'index': index}
        return self.variables_files_dict

    def add_variables_file(self, variables_file: str):
        self.fetch_variables_files()
        if os.path.isfile(variables_file):
            self.variables_files_com_obj.Add(variables_file)
            self.fetch_variables_files()
            self.__log.debug(f'Added the Variables file ({variables_file}) to the collection.')
        else:
            self.__log.warning(f'The given file ({variables_file}) does not exist.')

    def remove_variables_file(self, variables_file_name: str):
        self.fetch_variables_files()
        if variables_file_name in self.variables_files_dict:
            self.variables_files_com_obj.Remove(variables_file_name)
            self.fetch_variables_files()
            self.__log.debug(f'Removed the Variables file ({variables_file_name}) from the collection.')
        else:
            self.__log.warning(f'The given file ({variables_file_name}) does not exist.')

    def fetch_namespace_namespaces(self, parent_namespace_com_obj, parent_namespace_name):
        namespaces_count = parent_namespace_com_obj.Namespaces.Count
        if namespaces_count > 0:
            for index in range(1, namespaces_count + 1):
                namespace_com_obj = win32com.client.Dispatch(parent_namespace_com_obj.Namespaces.Item(index))
                namespace_name = f'{parent_namespace_name}::{namespace_com_obj.Name}'
                self.namespaces_dict[namespace_name] = namespace_com_obj
                if 'Namespaces' in dir(namespace_com_obj):
                    self.fetch_namespace_namespaces(namespace_com_obj, namespace_name)
                if 'Variables' in dir(namespace_com_obj):
                    self.fetch_namespace_variables(namespace_com_obj)

    def fetch_namespace_variables(self, parent_namespace_com_obj):
        variables_count = parent_namespace_com_obj.Variables.Count
        if variables_count > 0:
            for index in range(1, variables_count + 1):
                variable_obj = Variable(parent_namespace_com_obj.Variables.Item(index))
                self.variables_dict[variable_obj.full_name] = variable_obj

    def add_system_variable(self, namespace, variable, value):
        self.fetch_namespaces()
        if f'{namespace}::{variable}' in self.variables_dict.keys():
            self.__log.warning(f'The given variable ({variable}) already exists in the namespace ({namespace}).')
            return None
        else:
            self.add_namespace(namespace)
            return self.namespaces_dict[namespace].Variables.Add(variable, value)

    def remove_system_variable(self, namespace, variable):
        self.fetch_namespaces()
        if f'{namespace}::{variable}' not in self.variables_dict.keys():
            self.__log.warning(f'The given variable ({variable}) already removed in the namespace ({namespace}).')
            return None
        else:
            self.namespaces_dict[namespace].Variables.Remove(variable)


class Variable:
    def __init__(self, variable_com_obj):
        try:
            self.com_obj = win32com.client.Dispatch(variable_com_obj)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe Variable: {str(e)}')

    @property
    def analysis_only(self) -> bool:
        return self.com_obj.AnalysisOnly

    @analysis_only.setter
    def analysis_only(self, value: bool) -> None:
        self.com_obj.AnalysisOnly = value

    @property
    def bit_count(self) -> int:
        return self.com_obj.BitCount

    @property
    def comment(self) -> str:
        return self.com_obj.Comment

    @comment.setter
    def comment(self, text: str) -> None:
        self.com_obj.Comment = text

    @property
    def element_count(self) -> int:
        return self.com_obj.ElementCount

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @full_name.setter
    def full_name(self, full_name: str) -> None:
        self.com_obj.FullName = full_name

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def init_value(self) -> tuple[int, float, str]:
        return self.com_obj.InitValue

    @property
    def min_value(self) -> tuple[int, float, str]:
        return self.com_obj.MinValue

    @property
    def max_value(self) -> tuple[int, float, str]:
        return self.com_obj.MaxValue

    @property
    def is_array(self) -> bool:
        return self.com_obj.IsArray

    @property
    def is_signed(self) -> bool:
        return self.com_obj.IsSigned

    @property
    def read_only(self) -> bool:
        return self.com_obj.ReadOnly

    @property
    def type(self) -> int:
        return self.com_obj.Type

    @property
    def unit(self) -> str:
        return self.com_obj.Unit

    @property
    def value(self) -> tuple[int, float, str]:
        return self.com_obj.Value

    @value.setter
    def value(self, value: tuple[int, float, str]) -> None:
        self.com_obj.Value = value

    def get_member_phys_value(self, member_name: str):
        return self.com_obj.GetMemberPhysValue(member_name)

    def get_member_value(self, member_name: str):
        return self.com_obj.GetMemberValue(member_name)

    def get_symbolic_value_name(self, value: int):
        return self.com_obj.GetSymbolicValueName(value)

    def set_member_phys_value(self, member_name: str, value):
        return self.com_obj.setMemberPhysValue(member_name, value)

    def set_member_value(self, member_name: str, value):
        return self.com_obj.setMemberValue(member_name, value)

    def set_symbolic_value_name(self, value: int, name: str):
        self.com_obj.setSymbolicValueName(value, name)
