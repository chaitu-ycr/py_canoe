import win32com.client
from typing import Union


class SimulationEvents:
    EVENTS_INFORMATION = {}

    @staticmethod
    def OnIdle(timeHigh, time):
        SimulationEvents.EVENTS_INFORMATION['timeHigh'] = timeHigh
        SimulationEvents.EVENTS_INFORMATION['time'] = time


class Simulation:
    """
    The Simulation object represents CANoe's measurement functions in the Simulation mode.
    With the help of the Simulation object you can control the system time from an external source during the measurement.
    CANoe automatically goes into Slave mode at the measurement start if you access the Simulation object.
    """
    def __init__(self, app, enable_events: bool = False):
        self.com_object = app.com_object.Simulation
        if enable_events:
            win32com.client.WithEvents(self.com_object, SimulationEvents)

    @property
    def animation(self) -> Union[int, float]:
        return self.com_object.Animation

    @animation.setter
    def animation(self, value: Union[int, float]):
        self.com_object.Animation = value

    @property
    def current_time(self) ->int:
        return self.com_object.CurrentTime

    @property
    def current_time_high(self) -> int:
        return self.com_object.CurrentTimeHigh

    @property
    def notification_type(self) -> int:
        return self.com_object.NotificationType

    @notification_type.setter
    def notification_type(self, value: int):
        self.com_object.NotificationType = value

    def increment_time(self, ticks: int):
        self.com_object.IncrementTime(ticks)

    def increment_time_and_wait(self, ticks: int):
        self.com_object.IncrementTimeAndWait(ticks)
