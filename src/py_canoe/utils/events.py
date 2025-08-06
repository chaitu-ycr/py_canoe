import time
import pythoncom
from typing import Union
from datetime import datetime
from py_canoe.utils.common import logger
import py_canoe.utils.diagnostic


def wait(timeout_seconds=0.1):
    """Pump waiting COM messages and sleep for the given timeout."""
    pythoncom.PumpWaitingMessages()
    time.sleep(timeout_seconds)

def DoEvents() -> None:
    wait(0.05)

def DoEventsUntil(cond, timeout, title) -> bool:
    base_time = datetime.now()
    while not cond():
        DoEvents()
        now = datetime.now()
        difference = now - base_time
        seconds = difference.seconds
        if seconds > timeout:
            logger.warning(f'⚠️ {title} timeout({timeout} s)')
            return False
    return True


class ApplicationEvents:
    CONFIGURATION_OPENED: bool = False
    CANOE_IS_QUIT: bool = False

    @staticmethod
    def OnOpen(fullname):
        ApplicationEvents.CONFIGURATION_OPENED = True
        logger.info(f'[EVENT][APPLICATION] CANoe Configuration Opened: {fullname}')

    @staticmethod
    def OnQuit():
        ApplicationEvents.CANOE_IS_QUIT = True
        logger.info('[EVENT][APPLICATION] CANoe Application Quit')


def wait_for_event_canoe_configuration_to_open(timeout: Union[int, float]) -> bool:
    ApplicationEvents.CONFIGURATION_OPENED = False
    status = DoEventsUntil(lambda: ApplicationEvents.CONFIGURATION_OPENED, timeout, "CANoe Configuration Open")
    if not status:
        logger.error(f"😡 Error: CANoe configuration did not open within {timeout} seconds.")
    return status

def wait_for_event_canoe_quit(timeout: Union[int, float]) -> bool:
    ApplicationEvents.CANOE_IS_QUIT = False
    status = DoEventsUntil(lambda: ApplicationEvents.CANOE_IS_QUIT, timeout, "CANoe Application Quit")
    if not status:
        logger.error(f"😡 Error: CANoe application did not quit within {timeout} seconds.")
    return status


class MeasurementEvents:
    APP_COM_OBJ = object
    MEASUREMENT_IS_INITIALIZED: bool = False
    MEASUREMENT_IS_RUNNING: bool = False
    MEASUREMENT_IS_EXITED: bool = False
    CAPL_FUNCTION_NAMES = tuple()
    CAPL_FUNCTION_OBJECTS = dict()

    @staticmethod
    def OnInit():
        """measurement is initialized"""
        for fun in MeasurementEvents.CAPL_FUNCTION_NAMES:
            MeasurementEvents.CAPL_FUNCTION_OBJECTS[fun] = MeasurementEvents.APP_COM_OBJ.CAPL.GetFunction(fun)
        if MeasurementEvents.CAPL_FUNCTION_NAMES:
            logger.info(f'[EVENT][MEASUREMENT] Measurement Initialized with CAPL functions: {MeasurementEvents.CAPL_FUNCTION_NAMES}')
        else:
            MeasurementEvents.CAPL_FUNCTION_NAMES = tuple(MeasurementEvents.CAPL_FUNCTION_OBJECTS.keys())
            logger.info('[EVENT][MEASUREMENT] Measurement Initialized')
        MeasurementEvents.MEASUREMENT_IS_INITIALIZED = True

    @staticmethod
    def OnExit():
        """measurement is exited"""
        logger.info('[EVENT][MEASUREMENT] Measurement Exited')
        MeasurementEvents.CAPL_FUNCTION_OBJECTS.clear()
        MeasurementEvents.MEASUREMENT_IS_EXITED = True

    @staticmethod
    def OnStart():
        """measurement is started"""
        logger.info('[EVENT][MEASUREMENT] Measurement Started')
        MeasurementEvents.MEASUREMENT_IS_RUNNING = True

    @staticmethod
    def OnStop():
        """measurement is stopped"""
        logger.info('[EVENT][MEASUREMENT] Measurement Stopped')
        MeasurementEvents.MEASUREMENT_IS_RUNNING = False


def wait_for_event_canoe_measurement_started(timeout: Union[int, float], app_com_obj) -> bool:
    MeasurementEvents.APP_COM_OBJ = app_com_obj
    MeasurementEvents.MEASUREMENT_IS_INITIALIZED = False
    MeasurementEvents.MEASUREMENT_IS_RUNNING = False
    DoEventsUntil(lambda: MeasurementEvents.MEASUREMENT_IS_INITIALIZED, timeout, "CANoe Measurement Initialization")
    start_status = DoEventsUntil(lambda: MeasurementEvents.MEASUREMENT_IS_RUNNING, timeout, "CANoe Measurement Start")
    if not start_status:
        logger.error(f"😡 Error: CANoe measurement did not start within {timeout} seconds.")
    return start_status

def wait_for_event_canoe_measurement_stopped(timeout: Union[int, float], app_com_obj) -> bool:
    MeasurementEvents.APP_COM_OBJ = app_com_obj
    MeasurementEvents.MEASUREMENT_IS_RUNNING = True
    MeasurementEvents.MEASUREMENT_IS_EXITED = False
    stop_status = DoEventsUntil(lambda: not MeasurementEvents.MEASUREMENT_IS_RUNNING, timeout, "CANoe Measurement Stop")
    DoEventsUntil(lambda: MeasurementEvents.MEASUREMENT_IS_EXITED, timeout, "CANoe Measurement Exit")
    if not stop_status:
        logger.error(f"😡 Error: CANoe measurement did not stop within {timeout} seconds.")
    return stop_status


class ConfigurationEvents:
    CONFIGURATION_CLOSED: bool = False
    SYSTEM_VARIABLES_DEFINITION_CHANGED: bool = False

    @staticmethod
    def OnClose():
        logger.info('[EVENT][CONFIGURATION] CANoe Configuration Closed')
        ConfigurationEvents.CONFIGURATION_CLOSED = True

    @staticmethod
    def OnSystemVariablesDefinitionChanged():
        logger.info('[EVENT][CONFIGURATION] CANoe System Variables Definition Changed')
        ConfigurationEvents.SYSTEM_VARIABLES_DEFINITION_CHANGED = True


def wait_for_event_canoe_configuration_closed(timeout: Union[int, float]) -> bool:
    ConfigurationEvents.CONFIGURATION_CLOSED = False
    status = DoEventsUntil(lambda: ConfigurationEvents.CONFIGURATION_CLOSED, timeout, "CANoe Configuration Close")
    if not status:
        logger.error(f"😡 Error: CANoe configuration did not close within {timeout} seconds.")
    return status

def wait_for_event_canoe_system_variables_definition_changed(timeout: Union[int, float]) -> bool:
    ConfigurationEvents.SYSTEM_VARIABLES_DEFINITION_CHANGED = False
    status = DoEventsUntil(lambda: ConfigurationEvents.SYSTEM_VARIABLES_DEFINITION_CHANGED, timeout, "CANoe System Variables Definition Change")
    if not status:
        logger.error(f"😡 Error: CANoe system variables definition did not change within {timeout} seconds.")
    return status


class DiagnosticRequestEvents:
    TIMEOUT = False
    RECEIVED_RESPONSE = False
    RESPONSE: Union['py_canoe.utils.diagnostic.DiagnosticResponse', None] = None

    @staticmethod
    def OnCompletion():
        DiagnosticRequestEvents.TIMEOUT = False
        DiagnosticRequestEvents.RECEIVED_RESPONSE = False
        DiagnosticRequestEvents.RESPONSE = None

    @staticmethod
    def OnConfirmation():
        DiagnosticRequestEvents.TIMEOUT = False
        DiagnosticRequestEvents.RECEIVED_RESPONSE = False
        DiagnosticRequestEvents.RESPONSE = None


    @staticmethod
    def OnResponse(response):
        DiagnosticRequestEvents.TIMEOUT = False
        DiagnosticRequestEvents.RECEIVED_RESPONSE = True
        DiagnosticRequestEvents.RESPONSE = py_canoe.utils.diagnostic.DiagnosticResponse(response)

    @staticmethod
    def OnTimeout():
        DiagnosticRequestEvents.TIMEOUT = True
        DiagnosticRequestEvents.RECEIVED_RESPONSE = False
        DiagnosticRequestEvents.RESPONSE = None