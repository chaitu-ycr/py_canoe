# TODO: complete implementation of the Simulation class
import logging
import win32com.client

logging.getLogger('py_canoe')

class Simulation:
    def __init__(self, app):
        self.com_object = win32com.client.Dispatch(app.com_object.Simulation)

    @property
    def animation(self):
        return self.com_object.Animation

    @animation.setter
    def animation(self, value: bool):
        self.com_object.Animation = value

    @property
    def current_time(self):
        return self.com_object.CurrentTime

    @property
    def current_time_high(self):
        return self.com_object.CurrentTimeHigh

    @property
    def notification_type(self, value: int = 0):
        self.com_object.NotificationType = value

    def increment_time(self, ticks: int):
        self.com_object.IncrementTime(ticks)

    def increment_time_and_wait(self, ticks: int):
        self.com_object.IncrementTimeAndWait(ticks)
