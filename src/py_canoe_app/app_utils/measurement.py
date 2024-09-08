# import external modules here
import logging
import pythoncom
import win32com.client
from time import sleep as wait
from datetime import datetime

# import internal modules here

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
            logging.getLogger('CANOE_LOG').info(f'⌛ measurement event timeout({timeout()} s).')
            break

class CanoeMeasurementEvents:
    """Handler for CANoe Measurement events"""
    app_com_obj = object
    user_capl_function_names = tuple()
    user_capl_function_obj_dict = dict()

    @staticmethod
    def OnInit():
        """Occurs when the measurement is initialized."""
        app_com_obj_loc = CanoeMeasurementEvents.app_com_obj
        for fun in CanoeMeasurementEvents.user_capl_function_names:
            CanoeMeasurementEvents.user_capl_function_obj_dict[fun] = app_com_obj_loc.CAPL.GetFunction(fun)
        Measurement.STARTED = False
        Measurement.STOPPED = False


    @staticmethod
    def OnStart():
        """Occurs when the measurement is started."""
        Measurement.STARTED = True
        Measurement.STOPPED = False


    @staticmethod
    def OnStop():
        """Occurs when the measurement is stopped."""
        Measurement.STARTED = False
        Measurement.STOPPED = True

    @staticmethod
    def OnExit():
        """Occurs when the measurement is exited."""
        Measurement.STARTED = False
        Measurement.STOPPED = False

class Measurement:
    """The Measurement object represents measurement functions of CANoe."""
    STARTED = False
    STOPPED = False
    def __init__(self, app_com_obj, user_capl_function_names=tuple(), enable_meas_events=True):
        """The Measurement object represents measurement functions of CANoe."""
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
        """Returns the animation delay during the measurement in offline mode."""
        return self.com_obj.AnimationDelay

    @animation_delay.setter
    def animation_delay(self, delay: int):
        """Sets the animation delay during the measurement in offline mode.

        Args:
            delay (int): The animation delay.
        """
        self.com_obj.AnimationDelay = delay
        self.__log.info(f"Animation delay set to {delay}.")

    @property
    def measurement_index(self) -> int:
        """Returns the measurement index for the next measurement."""
        return self.com_obj.MeasurementIndex

    @measurement_index.setter
    def measurement_index(self, index: int):
        """Sets the measurement index for the next measurement.

        Args:
            index (int): The index of the measurement.
        """
        self.com_obj.MeasurementIndex = index
        self.__log.info(f"Measurement index set to {index}.")

    @property
    def running(self) -> bool:
        """Returns the running state of the measurement."""
        return self.com_obj.Running

    @property
    def user_capl_function_obj_dict(self):
        return CanoeMeasurementEvents.user_capl_function_obj_dict

    def animate(self):
        """Starts the measurement in animation mode."""
        self.com_obj.Animate()

    def break_offline_mode(self):
        """Interrupts the playback in offline mode."""
        self.com_obj.Break()

    def reset_offline_mode(self):
        """Resets the measurement in offline mode."""
        self.com_obj.Reset()

    def start(self):
        """Starts the measurement."""
        self.com_obj.Start()

    def step(self):
        """Processes a measurement event in single step."""
        self.com_obj.Step()

    def stop(self):
        self.stop_ex()

    def stop_ex(self):
        """StopEx repairs differences in the behavior of the deprecatedStop method (for the Measurement object) on deferred stops concerning simulated and real mode in CANoe.
        Calling the StopEx method correlates to clicking the Stop button (Home ribbon tab).
        The function must not be called if measurement has already ended.
        """
        self.com_obj.StopEx()