# TODO: Implement Capl
import logging
import win32com.client

logging.getLogger('py_canoe')

class Capl:
    """
    The CAPL object allows to compile all nodes (CAPL, .NET, XML) in the configuration. Additionally it represents the CAPL functions available in the CAPL programs.
    Please note that only user-defined CAPL functions can be accessed.
    """
    def __init__(self, app):
        self.com_object = win32com.client.Dispatch(app.com_object.CAPL)

    @property
    def compile_result(self) -> 'CompileResult':
        return CompileResult(self)

    def compile(self):
        self.com_object.Compile()

    def get_function(self, name: str) -> 'CaplFunction':
        return CaplFunction(self, name)

class CompileResult:
    """
    The CompileResult object represents the result of the last compilation of the CAPL object.
    """
    def __init__(self, capl):
        self.com_object = win32com.client.Dispatch(capl.com_object.CompileResult)

    @property
    def error_message(self) -> str:
        return self.com_object.ErrorMessage

    @property
    def node_name(self) -> str:
        return self.com_object.NodeName

    @property
    def result(self) -> int:
        return self.com_object.result

    @property
    def source_file(self) -> str:
        return self.com_object.SourceFile

class CaplFunction:
    """
    The CAPLFunction object represents a CAPL function.
    """
    def __init__(self, capl, name: str):
        self.com_object = win32com.client.Dispatch(capl.com_object.GetFunction(name))

    @property
    def parameter_count(self) -> int:
        return self.com_object.ParameterCount

    @property
    def parameter_types(self) -> list:
        return self.com_object.ParameterTypes

    def call(self, *parameters):
        return self.com_object.Call(*parameters)