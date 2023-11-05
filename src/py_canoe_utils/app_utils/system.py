# Import Python Libraries here
import logging
import win32com.client


class System:
    def __init__(self, app_com_obj: object):
        self.log = logging.getLogger('CANOE_LOG')
        self.com_obj = win32com.client.Dispatch(app_com_obj.System)


class VariablesFiles:
    def __init__(self):
        pass


class VariablesFile:
    def __init__(self):
        pass


class Namespaces:
    def __init__(self, namespaces_com_obj: object):
        self.namespaces_com_obj = namespaces_com_obj

    @property
    def count(self) -> int:
        return self.namespaces_com_obj.Count

    def add(self, name: str) -> object:
        """Adds a new namespace.

        Args:
            name (str): The name of the new namespace.

        Returns:
            object: The new Namespace object.
        """
        return self.namespaces_com_obj.Add(name)

    def remove(self, name: str) -> None:
        """Removes an Namespace from a group

        Args:
            name (str): A Namespace object.
        """
        self.namespaces_com_obj.Remove(name)


class Namespace:
    def __init__(self, namespace_com_obj: object):
        self.namespace_com_obj = namespace_com_obj

    @property
    def comment(self) -> str:
        """Gets the comment for the Namespace.

        Returns:
            str: The comment.
        """
        return self.namespace_com_obj.Comment

    @comment.setter
    def comment(self, text: str) -> None:
        """Defines the comment for the Namespace.

        Args:
            text (str): The comment
        """
        self.namespace_com_obj.Comment = text

    @property
    def name(self) -> str:
        """Returns the name of the Namespace.

        Returns:
            str: The name of the namespace.
        """
        return self.namespace_com_obj.Name

    @property
    def namespaces(self) -> object:
        """Returns the Namespaces object.

        Returns:
            object: The Namespaces object.
        """
        return self.namespace_com_obj.Namespaces

    @property
    def variables(self) -> object:
        """Returns the Variables object.

        Returns:
            object: The Variables object.
        """
        return self.namespace_com_obj.Variables


class Variables:
    def __init__(self, variables_com_obj):
        self.variables_com_obj = variables_com_obj

    @property
    def count(self) -> int:
        """Returns the number of Variable objects inside the collection.

        Returns:
            int: _description_
        """
        return self.variables_com_obj.Count

    def add(self, name: str, initial_value=0) -> object:
        """Adds a new read-only variable.

        Args:
            name (str): The name of the new variable.
            initial_value (int): The initial value of the new variable. default value: 0 (Integer).

        Returns:
            object: The new Variable object.
        """
        return self.variables_com_obj.Add(name, initial_value)

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
        return self.variables_com_obj.AddEx(name, initial_value, min_value, max_value)

    def add_writable(self, name: str, initial_value=0) -> object:
        """Adds a new writable variable.

        Args:
            name (str): The name of the new variable.
            initial_value (int): The initial value of the new variable. default value: 0 (Integer).

        Returns:
            object: The new Variable object.
        """
        return self.variables_com_obj.AddWriteable(name, initial_value)

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
        return self.variables_com_obj.AddWritableEx(name, initial_value, min_value, max_value)

    def remove(self, variable: object) -> None:
        """Removes variable from a group

        Args:
            variable (str): Variable object.
        """
        self.variables_com_obj.Remove(variable)


class Variable:
    def __init__(self, variable_com_obj):
        self.variable_com_obj = variable_com_obj

    @property
    def analysis_only(self) -> bool:
        """Determines if the variable shall be only used for analysis purposes or not.

        Returns:
            bool: false (default) ,true
        """
        return self.variable_com_obj.AnalysisOnly

    @analysis_only.setter
    def analysis_only(self, value: bool) -> None:
        """sets if the variable shall be only used for analysis purposes or not.
        If the property is set to false (default value), it may still be changed to analysis only in a CAPL program.

        Args:
            value (bool): false (default) ,true.
        """
        self.variable_com_obj.AnalysisOnly = value

    @property
    def bit_count(self) -> int:
        """Returns the number of bits of the variable data type.

        Returns:
            int: The number of bits of the variable data type.
        """
        return self.variable_com_obj.BitCount

    @property
    def comment(self) -> str:
        """Gets the comment for the variable..

        Returns:
            str: The comment.
        """
        return self.variable_com_obj.Comment

    @comment.setter
    def comment(self, text: str) -> None:
        """defines the comment for the variable.

        Args:
            text (str): The comment
        """
        self.variable_com_obj.Comment = text

    @property
    def element_count(self) -> int:
        """For arrays: the maximum number of elements in the array.

        Returns:
            int: The maximum number of elements in the array.
        """
        return self.variable_com_obj.ElementCount

    @property
    def full_name(self) -> str:
        """determines the complete path of the variable.

        Returns:
            str: The full name, including namespace, variable name and member name.
        """
        return self.variable_com_obj.FullName

    @full_name.setter
    def full_name(self, full_name: str) -> None:
        """sets the complete path of the variable.

        Args:
            full_name (str): The new complete path of the object.
        """
        self.variable_com_obj.FullName = full_name

    @property
    def name(self) -> str:
        """Returns the name of the variable.

        Returns:
            str: The name of the system variable.
        """
        return self.variable_com_obj.Name

    @property
    def init_value(self) -> tuple[int, float, str]:
        """The initial value of the variable.

        Returns:
            tuple[int, float, str]: The initial value of the variable.
        """
        return self.variable_com_obj.InitValue

    @property
    def min_value(self) -> tuple[int, float, str]:
        """Returns the minimum value of the object.

        Returns:
            tuple[int, float, str]: minimum value of the variable.
        """
        return self.variable_com_obj.MinValue

    @property
    def max_value(self) -> tuple[int, float, str]:
        """Returns the maximum value of the variable.

        Returns:
            tuple[int, float, str]: The maximum value of the variable.
        """
        return self.variable_com_obj.MaxValue

    @property
    def is_array(self) -> bool:
        """Whether the variable data type is an array.

        Returns:
            bool: Whether the variable data type is an array.
        """
        return self.variable_com_obj.IsArray

    @property
    def is_signed(self) -> bool:
        """For integer variables: whether the data type is signed.

        Returns:
            bool: Whether the data type is signed.
        """
        return self.variable_com_obj.IsSigned

    @property
    def read_only(self) -> bool:
        """Indicates whether the system variable is write protected.

        Returns:
            bool: If the variable is write protected True is returned; otherwise False is returned.
        """
        return self.variable_com_obj.ReadOnly

    @property
    def type(self) -> int:
        """Returns the type of a system variable.

        Returns:
            int: The type of the system variable. The following types are define- 0: Integer 1: Float 2: String 4: Float Array 5: Integer Array 6: LongLong 7: Byte Array 98: Generic Array 99: Struct 65535: Invalid
        """
        return self.variable_com_obj.Type

    @property
    def unit(self) -> str:
        """Returns the unit of the variable.

        Returns:
            str: The unit of the variable.
        """
        return self.variable_com_obj.Unit

    @property
    def value(self) -> tuple[int, float, str]:
        """Defines or sets the active value of the variable.

        Returns:
            tuple[int, float, str]: The value of the variable.
        """
        return self.variable_com_obj.Value

    @value.setter
    def value(self, value: tuple[int, float, str]) -> None:
        """Defines or sets the active value of the variable.

        Args:
            value (tuple[int, float, str]): The new value of the variable.
        """
        self.variable_com_obj.Value = value


class Encodings:
    def __init__(self):
        pass


class Encoding:
    def __init__(self):
        pass
