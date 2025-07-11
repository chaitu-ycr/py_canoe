import logging
import win32com.client

logging.getLogger('py_canoe')

class Performance:
    def __init__(self, app):
        self.com_object = win32com.client.Dispatch(app.com_object.Performance)

    @property
    def max_num_measurement_setup_threads(self) -> int:
        return self.com_object.MaxNumMeasurementSetupThreads

    @max_num_measurement_setup_threads.setter
    def max_num_measurement_setup_threads(self, num: int) -> None:
        # TODO: implement this method to work only when measurement not running
        self.com_object.MaxNumMeasurementSetupThreads = num
