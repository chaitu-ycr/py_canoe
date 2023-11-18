# Import Python Libraries here
import logging
import win32com.client


class System:
    """The System object represents the system of the CANoe application.
    The System object offers access to the namespaces for data exchange with external applications.
    """
    def __init__(self, app_com_obj):
        self.__log = logging.getLogger('CANOE_LOG')
        self.com_obj = win32com.client.Dispatch(app_com_obj.System)
        self.__com_obj_dir = dir(self.com_obj)

    @property
    def namespaces(self):
        """Returns the Namespaces object.
        """
        return Namespaces(self.com_obj.Namespaces)

    @property
    def variables_files(self):
        """Returns the VariablesFiles object.
        """
        return VariablesFiles(self.com_obj.VariablesFiles)


class Namespaces:
    """The Namespaces class represents the namespaces of the CANoe application"""

    def __init__(self, namespaces_com_obj):
        self.com_obj = win32com.client.Dispatch(namespaces_com_obj)
        self.__com_obj_dir = dir(self.com_obj)

    @property
    def count(self) -> int:
        """The number of namespaces contained"""
        return self.com_obj.Count

    def add(self, name: str) -> object:
        """Adds a new namespace.

        Args:
            name (str): The name of the new namespace.

        Returns:
            object: The new Namespace object.
        """
        return self.com_obj.Add(name)

    def remove(self, name: str) -> None:
        """Removes an Namespace from a group

        Args:
            name (str): A Namespace object.
        """
        self.com_obj.Remove(name)

    def fetch_namespaces(self) -> dict:
        namespaces_data = dict()
        if self.count > 0:
            for index in range(1, self.count + 1):
                namespace_com_obj = self.com_obj.Item(index)
                namespace = Namespace(namespace_com_obj)
                namespaces_data[namespace.name] = namespace
        return namespaces_data


class Namespace:
    def __init__(self, namespace_com_obj):
        self.com_obj = win32com.client.Dispatch(namespace_com_obj)
        self.__com_obj_dir = dir(self.com_obj)

    @property
    def comment(self) -> str:
        """The comment for the namespace.

        Returns:
            str: The comment.
        """
        return self.com_obj.Comment

    @comment.setter
    def comment(self, text: str) -> None:
        """Defines the comment for the Namespace.

        Args:
            text (str): The comment
        """
        self.com_obj.Comment = text

    @property
    def name(self) -> str:
        """Returns the name of the Namespace.

        Returns:
            str: The name of the namespace.
        """
        return self.com_obj.Name

    @property
    def namespaces(self):
        """Returns the Namespaces object.
        """
        return Namespaces(self.com_obj.Namespaces) if 'Namespaces' in self.__com_obj_dir else None

    @property
    def variables(self):
        """Returns the Variables object.
        """
        return Variables(self.com_obj.Variables) if 'Variables' in self.__com_obj_dir else None


class Variables:
    def __init__(self, variables_com_obj):
        self.com_obj = win32com.client.Dispatch(variables_com_obj)
        self.__com_obj_dir = dir(self.com_obj)

    @property
    def count(self) -> int:
        """Returns the number of Variable objects inside the collection.

        Returns:
            int: _description_
        """
        return self.com_obj.Count

    def add(self, name: str, initial_value=0) -> object:
        """Adds a new read-only variable.

        Args:
            name (str): The name of the new variable.
            initial_value (int): The initial value of the new variable. default value: 0 (Integer).

        Returns:
            object: The new Variable object.
        """
        return self.com_obj.Add(name, initial_value)

    def add_ex(self, name: str, initial_value=0, min_value=0, max_value=0) -> object:
        """Adds a new read-only variable.

        Args:
            name (str): The name of the new variable.
            initial_value (int, optional): The initial value of the new variable. Defaults to 0.
            min_value (int, optional): The minimum value of the new variable. Defaults to 0.
            max_value (int, optional): The maximum value of the new variable. Defaults to 0.

        Returns:
            object: The new Variable object.
        """
        return self.com_obj.AddEx(name, initial_value, min_value, max_value)

    def add_writable(self, name: str, initial_value=0) -> object:
        """Adds a new writable variable.

        Args:
            name (str): The name of the new variable.
            initial_value (int): The initial value of the new variable. default value: 0 (Integer).

        Returns:
            object: The new Variable object.
        """
        return self.com_obj.AddWriteable(name, initial_value)

    def add_writable_ex(self, name: str, initial_value=0, min_value=0, max_value=0) -> object:
        """Adds a new writable variable.

        Args:
            name (str): The name of the new variable.
            initial_value (int, optional): The initial value of the new variable. Defaults to 0.
            min_value (int, optional): The minimum value of the new variable. Defaults to 0.
            max_value (int, optional): The maximum value of the new variable. Defaults to 0.

        Returns:
            object: The new Variable object.
        """
        return self.com_obj.AddWritableEx(name, initial_value, min_value, max_value)

    def remove(self, variable: object) -> None:
        """Removes variable from a group

        Args:
            variable (str): Variable object.
        """
        self.com_obj.Remove(variable)

    def fetch_variables(self):
        variables_data = dict()
        if self.count > 0:
            for index in range(1, self.count + 1):
                variable_com_obj = self.com_obj.Item(index)
                variable = Variable(variable_com_obj)
                variables_data[variable.name] = variable
        return variables_data


class Variable:
    def __init__(self, variable_com_obj):
        self.com_obj = win32com.client.Dispatch(variable_com_obj)
        self.__com_obj_dir = dir(self.com_obj)

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


class VariablesFiles:
    def __init__(self, variables_files_com_obj):
        self.com_obj = win32com.client.Dispatch(variables_files_com_obj)
        self.__com_obj_dir = dir(self.com_obj)

    @property
    def count(self):
        """returns the number of variables contained.
        """
        return self.com_obj.Count

    def fetch_variables_files(self):
        variables_files_dict = dict()
        if self.count > 0:
            for index in range(1, self.count + 1):
                variable_file_com_obj = self.com_obj.Item(index)
                variable_file = VariablesFile(variable_file_com_obj)
                variables_files_dict[variable_file.name] = variable_file
        return variables_files_dict


class VariablesFile:
    def __init__(self, variable_file_com_obj):
        self.com_obj = win32com.client.Dispatch(variable_file_com_obj)
        self.__com_obj_dir = dir(self.com_obj)

    @property
    def full_name(self):
        """The complete path of the system variables file."""
        return self.com_obj.FullName

    @property
    def name(self):
        """The name of the system variables file."""
        return self.com_obj.Name

    @property
    def path(self):
        """The complete path to the system variables file."""
        return self.com_obj.Path


class Encodings:
    def __init__(self, encodings_com_obj):
        self.com_obj = win32com.client.Dispatch(encodings_com_obj)
        self.__com_obj_dir = dir(self.com_obj)

    @property
    def count(self):
        """The number of encodings contained"""
        return self.com_obj.Count

    def fetch_encodings(self):
        encodings_list = list()
        if self.count > 0:
            for index in range(1, self.count + 1):
                encoding_com_obj = self.com_obj.Item(index)
                encoding = Encoding(encoding_com_obj)
                encodings_list.append(encoding)
        return encodings_list


class Encoding:
    """The linear or textual encodings of the variable."""

    def __init__(self, encoding_com_obj):
        self.com_obj = win32com.client.Dispatch(encoding_com_obj)
        self.__com_obj_dir = dir(self.com_obj)

    @property
    def factor(self):
        """The factor of a linear encoding.
        Type: double
        """
        return self.com_obj.Factor

    @property
    def lower_bound(self):
        """The lower bound of the encoding.
        Type: 64 bit integer
        """
        return self.com_obj.LowerBound

    @property
    def offset(self):
        """The offset of a linear encoding.
        Type: double
        """
        return self.com_obj.Offset

    @property
    def text(self):
        """The textual value of a textual encoding.
        Type: string
        """
        return self.com_obj.Text

    @property
    def unit(self):
        """The unit of a linear encoding.
        Type: string
        """
        return self.com_obj.Unit

    @property
    def upper_bound(self):
        """The upper bound of the encoding.
        Type: 64 bit integer
        """
        return self.com_obj.UpperBound
