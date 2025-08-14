from typing import Union
import win32com.client

from py_canoe.core.capl import CaplFunction
from py_canoe.utils.common import DoEventsUntil
from py_canoe.utils.common import logger


class MeasurementEvents:
    def __init__(self):
        self.APP_COM_OBJ = object
        self.INIT: bool = False
        self.START: bool = False
        self.STOP: bool = False
        self.EXIT: bool = False
        self.CAPL_FUNCTION_OBJECTS = dict()
        self.CAPL_FUNCTION_NAMES = tuple()

    def OnInit(self):
        """measurement is initialized"""
        for fun in self.CAPL_FUNCTION_NAMES:
            self.CAPL_FUNCTION_OBJECTS[fun] = CaplFunction(self.APP_COM_OBJ.CAPL.GetFunction(fun))
        self.INIT = True

    def OnStart(self):
        """measurement is started"""
        self.START = True

    def OnStop(self):
        """measurement is stopped"""
        self.STOP = True

    def OnExit(self):
        """measurement is exited"""
        self.CAPL_FUNCTION_OBJECTS.clear()
        self.EXIT = True


class Measurement:
    def __init__(self, app):
        self.com_object = win32com.client.Dispatch(app.com_object.Measurement)
        self.measurement_events: MeasurementEvents = win32com.client.WithEvents(self.com_object, MeasurementEvents)
        self.measurement_events.APP_COM_OBJ = app.com_object

    @property
    def animation_delay(self) -> int:
        return self.com_object.AnimationDelay

    @animation_delay.setter
    def animation_delay(self, delay: int):
        self.com_object.AnimationDelay = delay
        logger.info(f"📢 Animation Delay ⏲️ set to: {delay} ms")

    @property
    def measurement_index(self) -> int:
        index = self.com_object.MeasurementIndex
        logger.info(f"📢 Measurement Index value: {index}")
        return index

    @measurement_index.setter
    def measurement_index(self, index: int):
        self.com_object.MeasurementIndex = index
        logger.info(f"📢 Measurement Index set to: {index}")

    @property
    def running(self) -> bool:
        return self.com_object.Running

    def start(self, timeout=30) -> bool:
        try:
            if self.running:
                logger.warning("⚠️ Measurement is already running")
                return True
            self.measurement_events.START = False
            self.com_object.Start()
            status = DoEventsUntil(lambda: self.measurement_events.START, timeout, "CANoe Measurement Start")
            if status:
                logger.info('📢 Measurement Started 🏃‍➡️')
            return status
        except Exception as e:
            logger.error(f"❌ Error starting CANoe measurement: {e}")
            return False

    def stop(self, timeout=30) -> bool:
        return self.stop_ex(timeout)

    def stop_ex(self, timeout=30) -> bool:
        try:
            if not self.running:
                logger.warning("⚠️ Measurement is already stopped")
                return True
            self.measurement_events.STOP = False
            self.com_object.Stop()
            status = DoEventsUntil(lambda: self.measurement_events.STOP, timeout, "CANoe Measurement Stop")
            if status:
                logger.info('📢 Measurement Stopped 🧍')
            return status
        except Exception as e:
            logger.error(f"❌ Error stopping CANoe measurement: {e}")
            return False

    def start_measurement_in_animation_mode(self, animation_delay=100, timeout=30) -> bool:
        try:
            if self.running:
                logger.warning("⚠️ Measurement is already running, cannot animate")
                return False
            self.measurement_events.START = False
            self.animation_delay = animation_delay
            self.com_object.Animate()
            status = DoEventsUntil(lambda: self.measurement_events.START, timeout, "CANoe Measurement Animation Initialization")
            if status:
                logger.info(f'📢 Measurement started 🏃‍➡️ in Animation mode with animation delay ⏲️ {animation_delay} ms')
            else:
                logger.error(f"❌ Measurement did not start in Animation mode within {timeout} seconds")
            return status
        except Exception as e:
            logger.error(f"❌ Error starting CANoe measurement in animation mode: {e}")
            return False

    def break_measurement_in_offline_mode(self) -> bool:
        try:
            if not self.running:
                logger.warning("⚠️ Measurement is not running, cannot break")
                return False
            self.com_object.Break()
            logger.info('📢 Measurement break applied 🫷 in Offline mode')
            return True
        except Exception as e:
            logger.error(f"❌ Error breaking CANoe measurement in offline mode: {e}")
            return False

    def reset_measurement_in_offline_mode(self) -> bool:
        try:
            self.com_object.Reset()
            logger.info('📢 Measurement reset applied 🔁 in Offline mode')
            return True
        except Exception as e:
            logger.error(f"❌ Error resetting CANoe measurement in offline mode: {e}")
            return False

    def process_measurement_event_in_single_step(self) -> bool:
        try:
            self.com_object.Step()
            logger.info('📢 Processed a measurement event in single step 👣')
            return True
        except Exception as e:
            logger.error(f"❌ Error processing CANoe measurement event in single step: {e}")
            return False
