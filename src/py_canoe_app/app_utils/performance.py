# import external modules here
import logging
import win32com.client


class Performance:
    """The Performance object allows setting or returning parameters that influence the performance on multicore processors."""
    def __init__(self, app_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.Performance)
        except Exception as e:
            self.__log.error(f'😡 Error initializing Performance object: {str(e)}')

    @property
    def max_num_meas_setup_threads(self):
        return self.com_obj.MaxNumMeasurementSetupThreads

    @max_num_meas_setup_threads.setter
    def max_num_meas_setup_threads(self, num: int):
        self.com_obj.MaxNumMeasurementSetupThreads = num