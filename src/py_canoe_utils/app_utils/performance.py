# Import Python Libraries here
import win32com.client


class Performance:
    def __init__(self, app_com_obj):
        self.com_obj = win32com.client.Dispatch(app_com_obj.Performance)

    @property
    def max_num_meas_setup_threads(self):
        return self.com_obj.MaxNumMeasurementSetupThreads

    @max_num_meas_setup_threads.setter
    def max_num_meas_setup_threads(self, num):
        self.com_obj.MaxNumMeasurementSetupThreads = num
