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
        self.__log = logging.getLogger('CANOE_LOG')
        self.com_obj = win32com.client.Dispatch(app_com_obj.System)
        self.namespaces_com_obj = win32com.client.Dispatch(self.com_obj.Namespaces)
        self.variables_files_com_obj = win32com.client.Dispatch(self.com_obj.VariablesFiles)
        self.namespaces_dict = {}
        self.variables_files_dict = {}
        self.variables_dict = {}

    @property
    def namespaces_count(self) -> int:
        return self.namespaces_com_obj.Count

    def fetch_namespaces(self) -> list:
        if self.namespaces_count > 0:
            for index in range(1, self.namespaces_count + 1):
                namespace_com_obj = win32com.client.Dispatch(self.namespaces_com_obj.Item(index))
                namespace_name = namespace_com_obj.Name
                self.namespaces_dict[namespace_name] = namespace_com_obj
                if 'Namespaces' in dir(namespace_com_obj):
                    self.fetch_namespace_namespaces(namespace_com_obj, namespace_name)
                if 'Variables' in dir(namespace_com_obj):
                    self.fetch_namespace_variables(namespace_com_obj)
        return set(self.namespaces_dict)

    def add_namespace(self, name: str):
        self.fetch_namespaces()
        if name not in self.namespaces_dict.keys():
            namespace_com_obj = self.namespaces_com_obj.Add(name)
            self.__log.info(f'Added the new namespace ({name}).')
            return namespace_com_obj
        else:
            self.__log.warning(f'The given namespace ({name}) already exists.')
            return None

    def remove_namespace(self, name: str) -> None:
        self.fetch_namespaces()
        if name in self.namespaces_list:
            self.namespaces_com_obj.Remove(name)
            self.fetch_namespaces()
            self.__log.info(f'Removed the namespace ({name}) from the collection.')
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
            self.__log.info(f'Added the Variables file ({variables_file}) to the collection.')
        else:
            self.__log.warning(f'The given file ({variables_file}) does not exist.')

    def remove_variables_file(self, variables_file_name: str):
        self.fetch_variables_files()
        if variables_file_name in self.variables_files_dict:
            self.variables_files_com_obj.Remove(variables_file_name)
            self.fetch_variables_files()
            self.__log.info(f'Removed the Variables file ({variables_file_name}) from the collection.')
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
            self.namespaces_dict[namespace].Variables.Add(variable, value)

    def remove_system_variable(self, namespace, variable):
        self.fetch_namespaces()
        if f'{namespace}::{variable}' not in self.variables_dict.keys():
            self.__log.warning(f'The given variable ({variable}) already removed in the namespace ({namespace}).')
            return None
        else:
            self.namespaces_dict[namespace].Variables.Remove(variable)


class Variable:
    def __init__(self, variable_com_obj):
        self.com_obj = win32com.client.Dispatch(variable_com_obj)

    @property
    def analysis_only(self) -> bool:
        """Determines if the variable shall be only used for analysis purposes or not.

        Returns:
            bool: false (default) ,true
        """
        return self.com_obj.AnalysisOnly

    @analysis_only.setter
    def analysis_only(self, value: bool) -> None:
        """sets if the variable shall be only used for analysis purposes or not.
        If the property is set to false (default value), it may still be changed to analysis only in a CAPL program.

        Args:
            value (bool): false (default) ,true.
        """
        self.com_obj.AnalysisOnly = value

    @property
    def bit_count(self) -> int:
        """Returns the number of bits of the variable data type.

        Returns:
            int: The number of bits of the variable data type.
        """
        return self.com_obj.BitCount

    @property
    def comment(self) -> str:
        """Gets the comment for the variable..

        Returns:
            str: The comment.
        """
        return self.com_obj.Comment

    @comment.setter
    def comment(self, text: str) -> None:
        """defines the comment for the variable.

        Args:
            text (str): The comment
        """
        self.com_obj.Comment = text

    @property
    def element_count(self) -> int:
        """For arrays: the maximum number of elements in the array.

        Returns:
            int: The maximum number of elements in the array.
        """
        return self.com_obj.ElementCount

    @property
    def full_name(self) -> str:
        """determines the complete path of the variable.

        Returns:
            str: The full name, including namespace, variable name and member name.
        """
        return self.com_obj.FullName

    @full_name.setter
    def full_name(self, full_name: str) -> None:
        """sets the complete path of the variable.

        Args:
            full_name (str): The new complete path of the object.
        """
        self.com_obj.FullName = full_name

    @property
    def name(self) -> str:
        """Returns the name of the variable.

        Returns:
            str: The name of the system variable.
        """
        return self.com_obj.Name

    @property
    def init_value(self) -> tuple[int, float, str]:
        """The initial value of the variable.

        Returns:
            tuple[int, float, str]: The initial value of the variable.
        """
        return self.com_obj.InitValue

    @property
    def min_value(self) -> tuple[int, float, str]:
        """Returns the minimum value of the object.

        Returns:
            tuple[int, float, str]: minimum value of the variable.
        """
        return self.com_obj.MinValue

    @property
    def max_value(self) -> tuple[int, float, str]:
        """Returns the maximum value of the variable.

        Returns:
            tuple[int, float, str]: The maximum value of the variable.
        """
        return self.com_obj.MaxValue

    @property
    def is_array(self) -> bool:
        """Whether the variable data type is an array.

        Returns:
            bool: Whether the variable data type is an array.
        """
        return self.com_obj.IsArray

    @property
    def is_signed(self) -> bool:
        """For integer variables: whether the data type is signed.

        Returns:
            bool: Whether the data type is signed.
        """
        return self.com_obj.IsSigned

    @property
    def read_only(self) -> bool:
        """Indicates whether the system variable is write protected.

        Returns:
            bool: If the variable is write protected True is returned; otherwise False is returned.
        """
        return self.com_obj.ReadOnly

    @property
    def type(self) -> int:
        """Returns the type of a system variable.

        Returns:
            int: The type of the system variable. The following types are define- 0: Integer 1: Float 2: String 4: Float Array 5: Integer Array 6: LongLong 7: Byte Array 98: Generic Array 99: Struct 65535: Invalid
        """
        return self.com_obj.Type

    @property
    def unit(self) -> str:
        """Returns the unit of the variable.

        Returns:
            str: The unit of the variable.
        """
        return self.com_obj.Unit

    @property
    def value(self) -> tuple[int, float, str]:
        """Defines or sets the active value of the variable.

        Returns:
            tuple[int, float, str]: The value of the variable.
        """
        return self.com_obj.Value

    @value.setter
    def value(self, value: tuple[int, float, str]) -> None:
        """Defines or sets the active value of the variable.

        Args:
            value (tuple[int, float, str]): The new value of the variable.
        """
        self.com_obj.Value = value

    def get_member_phys_value(self, member_name: str):
        """The current physical value of the member."""
        return self.com_obj.GetMemberPhysValue(member_name)

    def get_member_value(self, member_name: str):
        """The current (raw) value of the member."""
        return self.com_obj.GetMemberValue(member_name)

    def get_symbolic_value_name(self, value: int):
        """Returns the symbolic name for the value.
        Symbolic value names can only be used with variables of type Integer.
        """
        return self.com_obj.GetSymbolicValueName(value)

    def set_member_phys_value(self, member_name: str, value):
        """Sets the physical value of a member of a variable of type Struct or Generic Array."""
        return self.com_obj.setMemberPhysValue(member_name, value)

    def set_member_value(self, member_name: str, value):
        """Sets the value of a member of a variable of type Struct or Generic Array."""
        return self.com_obj.setMemberValue(member_name, value)

    def set_symbolic_value_name(self, value: int, name: str):
        """Defines the symbolic name for the value.
        An existing name for the value is replaced.
        Symbolic value names can only be used with variables of type Integer.
        """
        self.com_obj.setSymbolicValueName(value, name)
