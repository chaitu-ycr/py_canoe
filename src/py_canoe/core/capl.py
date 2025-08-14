from typing import Union

from py_canoe.utils.common import logger
from py_canoe.utils.common import wait


class CompileResult:
    """
    The CompileResult object represents the result of the last compilation of the CAPL object.
    """
    def __init__(self, compile_result):
        self.com_object = compile_result

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
    def __init__(self, capl_function):
        self.com_object = capl_function

    @property
    def parameter_count(self) -> int:
        return self.com_object.ParameterCount

    @property
    def parameter_types(self) -> list:
        return self.com_object.ParameterTypes

    def call(self, *parameters):
        return self.com_object.Call(*parameters)


class Capl:
    """
    The CAPL object allows to compile all nodes (CAPL, .NET, XML) in the configuration. Additionally it represents the CAPL functions available in the CAPL programs.
    Please note that only user-defined CAPL functions can be accessed.
    """
    def __init__(self, app):
        self.com_object = app.com_object.CAPL
        self.capl_function_objects = lambda: app.measurement.measurement_events.CAPL_FUNCTION_OBJECTS

    @property
    def compile_result(self) -> CompileResult:
        return CompileResult(self.com_object.CompileResult)

    def compile(self, wait_time: Union[int, float] = 5) -> Union[CompileResult, None]:
        try:
            self.com_object.Compile()
            wait(wait_time)
            compile_result = self.compile_result
            logger.info(f'🧑‍💻 compiled all CAPL nodes. result={compile_result.result}')
            return compile_result
        except Exception as e:
            logger.error(f"❌ Error compiling CAPL nodes: {e}")
            return None

    def get_function(self, name: str) -> CaplFunction:
        if name in self.capl_function_objects():
            return self.capl_function_objects()[name]
        else:
            logger.warning(f'⚠️ CAPL function "{name}" not found/registered.')
            return None

    def call_capl_function(self, name: str, *arguments) -> bool:
        try:
            capl_function = self.get_function(name)
            if capl_function:
                if len(arguments) != capl_function.parameter_count:
                    logger.warning(f"⚠️ Not enough arguments provided for CAPL function '{name}'.")
                    return False
                else:
                    if len(arguments) > 0:
                        capl_function.call(*arguments)
                    else:
                        capl_function.call()
                    return True
            else:
                logger.warning(f"⚠️ CAPL function '{name}' not found.")
                return False
        except Exception as e:
            logger.error(f"❌ Error calling CAPL function '{name}': {e}")
            return False
