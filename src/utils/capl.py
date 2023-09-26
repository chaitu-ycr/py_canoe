# Import Python Libraries here
import win32com.client
from typing import Union

class Capl:
    """The CAPL object allows to compile all nodes (CAPL, .NET, XML) in the configuration.
    Additionally it represents the CAPL functions available in the CAPL programs.
    Please note that only user-defined CAPL functions can be accessed
    """
    def __init__(self, app_obj: object) -> None:
        self.app_obj = app_obj
        self.log = self.app_obj.log
        self.capl_obj = win32com.client.Dispatch(self.app_obj.app_com_obj.CAPL)
    
    def compile(self) -> None:
        """Translates all CAPL, XML and .NET nodes.
        """
        self.capl_obj.Compile()

    def get_function(self, name: str) -> object:
        """Returns a CAPLFunction object.

        Args:
            name (str): The name of the CAPL function.

        Returns:
            object: The CAPLFunction object.
        """
        return self.capl_obj.GetFunction(name)

    @staticmethod
    def parameter_count(capl_function_object: object) -> int:
        """Returns the number of parameters of the CAPL function.

        Args:
            capl_function_object (object): The CAPLFunction object.

        Returns:
            int: The number of parameters of the CAPL function.
        """
        return capl_function_object.ParameterCount
    

    @staticmethod
    def parameter_types(capl_function_object: object) -> tuple:
        """Returns the types of the parameters of the CAPL function as byte array.
        The parameter types are coded as follows:
        L: long (32 bit signed integer)
        D: dword (32 bit unsigned integer)
        F: double (64 bit floating point)

        Args:
            capl_function_object (object): The CAPLFunction object.

        Returns:
            tuple: The types of the parameters of the CAPL function as byte array.
        """
        return capl_function_object.ParameterTypes
    
    def call_capl_function(self, name: str, *arguments) -> int:
        """Calls a CAPL function.
        Please note that the number of parameters must agree with that of the CAPL function.
        The return value is only available for CAPL functions whose CAPL programs are configured in the Measurement Setup.
        Only integers are allowed as a return type.

        Args:
            name (str): The name of the CAPL function.
            arguments (tuple): Function parameters p1…p10 (optional).

        Returns:
            int: The return value of the CAPL function.
        """
        capl_function_obj = self.get_function(name)
        if len(arguments) == self.parameter_count(capl_function_obj):
            if len(arguments) > 0:
                function_return_value = capl_function_obj.Call(*arguments)
            else:
                function_return_value = capl_function_obj.Call()
        else:
            function_return_value = None
            print(fr'function arguments not matching with CAPL user function args.')
        return function_return_value
    
    def compile_result(self) -> dict:
        """The CompileResult object represents the result of the last compilation of the CAPL object.

        Returns:
            dict: returns dictionary of 'error_message', 'node_name', 'result', 'source_file'
        """
        return_values = dict()
        compile_result_obj = self.capl_obj.CompileResult
        # Returns the last compilation error for the CompileResult object or the last loading error/warning for the OpenConfigurationResult object
        return_values['error_message'] = compile_result_obj.ErrorMessage
        # Returns the name of the first compilation error node.
        return_values['node_name'] = compile_result_obj.NodeName
        # Returns the result of the last compilation of the CAPL object.
        return_values['result'] = compile_result_obj.Result
        # Returns the path of the program file where the first compile error occurred
        return_values['source_file'] = compile_result_obj.SourceFile
        return return_values




    