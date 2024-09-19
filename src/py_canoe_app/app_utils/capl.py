# import external modules here
import logging
import win32com.client


class Capl:
    """The CAPL object allows to compile all nodes (CAPL, .NET, XML) in the configuration.
    Additionally, it represents the CAPL functions available in the CAPL programs.
    Please note that only user-defined CAPL functions can be accessed
    """

    def __init__(self, app_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.CAPL)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CAPL object: {str(e)}')

    def compile(self) -> None:
        self.com_obj.Compile()

    def get_function(self, name: str) -> object:
        return self.com_obj.GetFunction(name)

    @staticmethod
    def parameter_count(capl_function_object: get_function) -> int:
        return capl_function_object.ParameterCount

    @staticmethod
    def parameter_types(capl_function_object: get_function) -> tuple:
        return capl_function_object.ParameterTypes

    def call_capl_function(self, capl_function_obj: get_function, *arguments) -> bool:
        return_value = False
        if len(arguments) == self.parameter_count(capl_function_obj):
            if len(arguments) > 0:
                capl_function_obj.Call(*arguments)
            else:
                capl_function_obj.Call()
            return_value = True
        else:
            self.__log.warning(fr'😇 function arguments not matching with CAPL user function args.')
        return return_value

    def compile_result(self) -> dict:
        return_values = dict()
        compile_result_obj = self.com_obj.CompileResult
        return_values['error_message'] = compile_result_obj.ErrorMessage
        return_values['node_name'] = compile_result_obj.NodeName
        return_values['result'] = compile_result_obj.result
        return_values['source_file'] = compile_result_obj.SourceFile
        return return_values
