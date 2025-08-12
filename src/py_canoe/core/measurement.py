from typing import Union

from py_canoe.core.capl import CaplFunction
from py_canoe.utils.common import DoEventsUntil
from py_canoe.utils.common import logger


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
            MeasurementEvents.CAPL_FUNCTION_OBJECTS[fun] = CaplFunction(MeasurementEvents.APP_COM_OBJ.CAPL.GetFunction(fun))
        if MeasurementEvents.CAPL_FUNCTION_NAMES:
            logger.info(f'[EVENT][MEASUREMENT] 📢 Measurement Initialized with CAPL functions: {MeasurementEvents.CAPL_FUNCTION_NAMES}')
        else:
            MeasurementEvents.CAPL_FUNCTION_NAMES = tuple(MeasurementEvents.CAPL_FUNCTION_OBJECTS.keys())
            logger.info('[EVENT][MEASUREMENT] 📢 Measurement Initialized')
        MeasurementEvents.MEASUREMENT_IS_INITIALIZED = True

    @staticmethod
    def OnExit():
        """measurement is exited"""
        logger.info('[EVENT][MEASUREMENT] 📢 Measurement Exited')
        MeasurementEvents.CAPL_FUNCTION_OBJECTS.clear()
        MeasurementEvents.MEASUREMENT_IS_EXITED = True

    @staticmethod
    def OnStart():
        """measurement is started"""
        logger.info('[EVENT][MEASUREMENT] 📢 Measurement Started')
        MeasurementEvents.MEASUREMENT_IS_RUNNING = True

    @staticmethod
    def OnStop():
        """measurement is stopped"""
        logger.info('[EVENT][MEASUREMENT] 📢 Measurement Stopped')
        MeasurementEvents.MEASUREMENT_IS_RUNNING = False


def wait_for_event_canoe_measurement_started(timeout: Union[int, float], app_com_obj) -> bool:
    MeasurementEvents.APP_COM_OBJ = app_com_obj
    MeasurementEvents.MEASUREMENT_IS_INITIALIZED = False
    MeasurementEvents.MEASUREMENT_IS_RUNNING = False
    DoEventsUntil(lambda: MeasurementEvents.MEASUREMENT_IS_INITIALIZED, timeout, "CANoe Measurement Initialization")
    start_status = DoEventsUntil(lambda: MeasurementEvents.MEASUREMENT_IS_RUNNING, timeout, "CANoe Measurement Start")
    if not start_status:
        logger.error(f"😡 CANoe measurement did not start within {timeout} seconds.")
    return start_status

def wait_for_event_canoe_measurement_stopped(timeout: Union[int, float], app_com_obj) -> bool:
    MeasurementEvents.APP_COM_OBJ = app_com_obj
    MeasurementEvents.MEASUREMENT_IS_RUNNING = True
    MeasurementEvents.MEASUREMENT_IS_EXITED = False
    stop_status = DoEventsUntil(lambda: not MeasurementEvents.MEASUREMENT_IS_RUNNING, timeout, "CANoe Measurement Stop")
    DoEventsUntil(lambda: MeasurementEvents.MEASUREMENT_IS_EXITED, timeout, "CANoe Measurement Exit")
    if not stop_status:
        logger.error(f"😡 CANoe measurement did not stop within {timeout} seconds")
    return stop_status

def start_measurement(app, timeout=30) -> bool:
    try:
        if app.com_object.Measurement.Running:
            logger.warning("⚠️ Measurement is already running")
            return True
        app.com_object.Measurement.Start()
        return wait_for_event_canoe_measurement_started(timeout, app.com_object)
    except Exception as e:
        logger.error(f"😡 Error starting CANoe measurement: {e}")
        return False

def stop_measurement(app, timeout=30) -> bool:
    return app.stop_ex_measurement(timeout)

def stop_ex_measurement(app, timeout=60) -> bool:
    try:
        if not app.com_object.Measurement.Running:
            logger.warning("⚠️ Measurement is already stopped")
            return True
        app.com_object.Measurement.StopEx()
        return wait_for_event_canoe_measurement_stopped(timeout, app.com_object)
    except Exception as e:
        logger.error(f"😡 Error stopping CANoe measurement with StopEx: {e}")
        return False

def reset_measurement(app, timeout=30) -> bool:
    try:
        if not app.stop_ex_measurement(timeout=timeout):
            logger.error("😡 Error stopping measurement during reset")
            return False
        if not app.start_measurement(timeout=timeout):
            logger.error("😡 Error starting measurement during reset")
            return False
        logger.info("📢 Measurement reset 🔁 successfull 🎉")
        return True
    except Exception as e:
        logger.error(f"😡 Error resetting measurement: {e}")
        return False

def get_measurement_running_status(app) -> bool:
    try:
        return app.com_object.Measurement.Running
    except Exception as e:
        logger.error(f"😡 Error getting measurement running status: {e}")
        return False

def start_measurement_in_animation_mode(app, animation_delay=100, timeout=30) -> bool:
    try:
        if app.com_object.Measurement.Running:
            logger.info("Measurement is already running.")
            return True
        app.com_object.Measurement.AnimationDelay = animation_delay
        app.com_object.Measurement.Animate()
        started = wait_for_event_canoe_measurement_started(timeout, app.com_object)
        if started:
            logger.info(f'📢 Started 🏃‍♂️ measurement in Animation mode with animation delay ⏲️ {animation_delay} ms')
        return started
    except Exception as e:
        logger.error(f"😡 Error starting CANoe measurement in animation mode: {e}")
        return False

def break_measurement_in_offline_mode(app) -> bool:
    try:
        if not app.com_object.Measurement.Running:
            logger.info("Measurement is not running, cannot break.")
            return False
        app.com_object.Measurement.Break()
        logger.info('📢 Measurement break applied 🫷 in Offline mode')
        return True
    except Exception as e:
        logger.error(f"😡 Error breaking CANoe measurement in offline mode: {e}")
        return False

def reset_measurement_in_offline_mode(app) -> bool:
    try:
        app.com_object.Measurement.Reset()
        logger.info('📢 Measurement reset 🔁 in Offline mode')
        return True
    except Exception as e:
        logger.error(f"😡 Error resetting CANoe measurement in offline mode: {e}")
        return False

def step_measurement_event_in_single_step(app) -> bool:
    try:
        app.com_object.Measurement.Step()
        logger.info('📢 Processed a measurement event in single step 👣')
        return True
    except Exception as e:
        logger.error(f"😡 Error stepping CANoe measurement in single step mode: {e}")
        return False

def get_measurement_index(app) -> int:
    try:
        index = app.com_object.Measurement.MeasurementIndex
        logger.info(f"📢 Measurement Index retrieved: {index}")
        return index
    except Exception as e:
        logger.error(f"😡 Error retrieving CANoe measurement index: {e}")
        return -1

def set_measurement_index(app, index: int) -> bool:
    try:
        app.com_object.Measurement.MeasurementIndex = index
        logger.info(f"📢 Measurement Index set to: {index}")
        return True
    except Exception as e:
        logger.error(f"😡 Error setting CANoe measurement index: {e}")
        return False
