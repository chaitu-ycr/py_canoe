# Import Python Libraries here
import pythoncom
import win32com.client
from time import sleep as wait


class Measurement:
    """The Measurement object represents measurement functions of CANoe.
    """
    STARTED = False
    STOPPED = False

    def __init__(self, app, user_capl_function_names=tuple()) -> None:
        self.app = app
        self.log = self.app.log
        CanoeMeasurementEvents.app_com_obj = self.app.app_com_obj
        CanoeMeasurementEvents.user_capl_function_names = user_capl_function_names
        self.meas_obj = win32com.client.Dispatch(self.app.app_com_obj.Measurement)
        self.wait_for_canoe_meas_to_start = lambda: DoEventsUntil(lambda: Measurement.STARTED)
        self.wait_for_canoe_meas_to_stop = lambda: DoEventsUntil(lambda: Measurement.STOPPED)
        win32com.client.WithEvents(self.meas_obj, CanoeMeasurementEvents)

    @property
    def animation_delay(self) -> int:
        """Defines the animation delay during the measurement in Offline mode.

        Returns:
            int: The animation delay during the measurement in Offline mode.
        """
        return self.meas_obj.AnimationDelay

    @animation_delay.setter
    def animation_delay(self, delay: int):
        """Sets the animation delay during the measurement in Offline mode.

        Args:
            delay (int): Animation delay
        """
        self.meas_obj.AnimationDelay = delay
        self.log.info(f'Animation delay set to = {delay}.')

    @property
    def measurement_index(self) -> int:
        """Determines the measurement index for the next measurement.

        Returns:
            int: Returns the measurement index for the next measurement.
        """
        return self.meas_obj.MeasurementIndex

    @measurement_index.setter
    def measurement_index(self, index: int):
        """sets the measurement index for the next measurement.

        Args:
            index (int): The measurement index for the next measurement.
        """
        self.meas_obj.MeasurementIndex = index
        self.log.info(f'next measurement index set to = {index}.')

    @property
    def running(self) -> bool:
        """Returns the running state of the measurement.

        Returns:
            bool: True- The measurement is running. False- The measurement is not running.
        """
        return self.meas_obj.Running

    @property
    def user_capl_function_obj_dict(self):
        return CanoeMeasurementEvents.user_capl_function_obj_dict

    def animate(self) -> None:
        """Starts the measurement in Animation mode.
        """
        self.meas_obj.Animate()
        self.log.info(f'Started the measurement in Animation mode with animation delay = {self.animation_delay}.')

    def break_offline_mode(self) -> None:
        """Interrupts the playback in Offline mode.
        """
        if self.running:
            self.meas_obj.Break()
            self.log.info('Interrupted the playback in Offline mode.')

    def reset_offline_mode(self) -> None:
        """Resets the measurement in Offline mode.
        """
        self.meas_obj.Reset()
        self.log.info('resetted measurement in offline mode.')

    def start(self) -> bool:
        """Starts the measurement.
        """
        if not self.running:
            self.meas_obj.Start()
            if not self.running:
                self.log.info(f'waiting for measurement to start...')
                self.wait_for_canoe_meas_to_start()
            self.log.info(f'CANoe Measurement Started. Measurement running status = {self.running}')
        else:
            self.log.info(f'CANoe Measurement Already Running. Measurement running status = {self.running}')
        return self.running

    def step(self) -> None:
        """Processes a measurement event in single step.
        """
        if not self.running:
            self.meas_obj.Step()
            self.log.info('processed a measurement event in single step')

    def stop(self) -> bool:
        """Stops the measurement.
        """
        return self.stop_ex()

    def stop_ex(self) -> bool:
        """StopEx repairs differences in the behavior of the Stop method on deferred stops concerning simulated and real mode in CANoe.
        Calling the StopEx method correlates to clicking the Stop button.
        """
        if self.running:
            self.meas_obj.StopEx()
            if self.running:
                self.log.info(f'waiting for measurement to stop...')
                self.wait_for_canoe_meas_to_stop()
            self.log.info(f'CANoe Measurement Stopped. Measurement running status = {self.running}')
        else:
            self.log.info(f'CANoe Measurement Already Stopped. Measurement running status = {self.running}')
        return not self.running


class CanoeMeasurementEvents:
    """Handler for CANoe Measurement events"""
    app_com_obj = object
    user_capl_function_names = tuple()
    user_capl_function_obj_dict = dict()

    @staticmethod
    def OnInit():
        """Occurs when the measurement is initialized.
        """
        app_com_obj_loc = CanoeMeasurementEvents.app_com_obj
        for fun in CanoeMeasurementEvents.user_capl_function_names:
            CanoeMeasurementEvents.user_capl_function_obj_dict[fun] = app_com_obj_loc.CAPL.GetFunction(fun)
        print('measurement OnInit event triggered')

    @staticmethod
    def OnExit():
        """Occurs when the measurement is exited.
        """
        print('measurement OnExit event triggered')

    @staticmethod
    def OnStart():
        """Occurs when the measurement is started.
        """
        Measurement.STARTED = True
        Measurement.STOPPED = False
        print('measurement OnStart event triggered')

    @staticmethod
    def OnStop():
        """Occurs when the measurement is stopped.
        """
        Measurement.STARTED = False
        Measurement.STOPPED = True
        print('measurement OnStop event triggered')


def DoEvents():
    pythoncom.PumpWaitingMessages()
    wait(.1)


def DoEventsUntil(cond):
    while not cond():
        DoEvents()
