# import external modules here
import logging
import pythoncom
import win32com.client
from datetime import datetime
from time import sleep as wait


def DoEvents():
    pythoncom.PumpWaitingMessages()
    wait(.1)

def DoEventsUntil(cond, timeout):
    base_time = datetime.now()
    while not cond():
        DoEvents()
        now = datetime.now()
        difference = now - base_time
        seconds = difference.seconds
        if seconds > timeout():
            logging.getLogger('CANOE_LOG').debug(f'⌛ measurement event timeout({timeout()} s).')
            break


class CanoeMeasurementEvents:
    """Handler for CANoe Measurement events"""
    app_com_obj = object
    user_capl_function_names = tuple()
    user_capl_function_obj_dict = dict()

    @staticmethod
    def OnInit():
        app_com_obj_loc = CanoeMeasurementEvents.app_com_obj
        for fun in CanoeMeasurementEvents.user_capl_function_names:
            CanoeMeasurementEvents.user_capl_function_obj_dict[fun] = app_com_obj_loc.CAPL.GetFunction(fun)
        Measurement.STARTED = False
        Measurement.STOPPED = False


    @staticmethod
    def OnStart():
        Measurement.STARTED = True
        Measurement.STOPPED = False


    @staticmethod
    def OnStop():
        Measurement.STARTED = False
        Measurement.STOPPED = True

    @staticmethod
    def OnExit():
        Measurement.STARTED = False
        Measurement.STOPPED = False


class Measurement:
    """The Measurement object represents measurement functions of CANoe."""
    STARTED = False
    STOPPED = False
    def __init__(self, app_com_obj, user_capl_function_names=tuple(), enable_meas_events=True):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            CanoeMeasurementEvents.app_com_obj = app_com_obj
            CanoeMeasurementEvents.user_capl_function_names = user_capl_function_names
            self.com_obj = win32com.client.Dispatch(app_com_obj.Measurement)
            self.meas_start_stop_timeout = 60   # default value set to 60 seconds (1 minute)
            self.wait_for_canoe_meas_to_start = lambda: DoEventsUntil(lambda: Measurement.STARTED, lambda: self.meas_start_stop_timeout)
            self.wait_for_canoe_meas_to_stop = lambda: DoEventsUntil(lambda: Measurement.STOPPED, lambda: self.meas_start_stop_timeout)
            if enable_meas_events:
                win32com.client.WithEvents(self.com_obj, CanoeMeasurementEvents)
        except Exception as e:
            self.__log.error(f"Error in Measurement class: {str(e)}")

    @property
    def animation_delay(self) -> int:
        return self.com_obj.AnimationDelay

    @animation_delay.setter
    def animation_delay(self, delay: int):
        self.com_obj.AnimationDelay = delay

    @property
    def measurement_index(self) -> int:
        return self.com_obj.MeasurementIndex

    @measurement_index.setter
    def measurement_index(self, index: int):
        self.com_obj.MeasurementIndex = index

    @property
    def running(self) -> bool:
        return self.com_obj.Running

    @property
    def user_capl_function_obj_dict(self):
        return CanoeMeasurementEvents.user_capl_function_obj_dict

    def animate(self):
        self.com_obj.Animate()

    def break_offline_mode(self):
        self.com_obj.Break()

    def reset_offline_mode(self):
        self.com_obj.Reset()

    def start(self):
        self.com_obj.Start()

    def step(self):
        self.com_obj.Step()

    def stop(self):
        self.stop_ex()

    def stop_ex(self):
        self.com_obj.StopEx()