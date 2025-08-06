from typing import Union
from py_canoe.utils.common import logger
from py_canoe.utils.common import DoEventsUntil


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
