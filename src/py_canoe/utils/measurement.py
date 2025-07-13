# TODO: complete implementation of the Measurement class
import logging
import win32com.client

logging.getLogger('py_canoe')

class MeasurementEvents:
    @staticmethod
    def OnExit():
        logging.info("Measurement Exit Triggered.")

    @staticmethod
    def OnInit():
        logging.info("Measurement Initialization Triggered.")

    @staticmethod
    def OnStart():
        logging.info("Measurement Start Triggered.")

    @staticmethod
    def OnStop():
        logging.info("Measurement Stop Triggered.")

class Measurement:
    def __init__(self, app, enable_events: bool = True):
        self.com_object = win32com.client.Dispatch(app.com_object.Measurement)
        if enable_events:
            win32com.client.WithEvents(self.com_object, MeasurementEvents)

    @property
    def animation_delay(self) -> int:
        return self.com_object.AnimationDelay

    @animation_delay.setter
    def animation_delay(self, delay: int):
        self.com_object.AnimationDelay = delay

    @property
    def measurement_index(self) -> int:
        return self.com_object.MeasurementIndex

    @measurement_index.setter
    def measurement_index(self, index: int):
        self.com_object.MeasurementIndex = index

    @property
    def running(self) -> bool:
        return self.com_object.Running

    def animate(self):
        self.com_object.Animate()

    def break_(self):
        self.com_object.Break()

    def reset(self):
        self.com_object.Reset()

    def start(self):
        self.com_object.Start()

    def step(self):
        self.com_object.Step()

    def stop(self):
        self.com_object.Stop()

    def stop_ex(self):
        self.com_object.StopEx()
