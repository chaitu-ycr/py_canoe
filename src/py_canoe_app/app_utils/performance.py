# import external modules here
import logging
import win32com.client

# import internal modules here


class Performance:
    """The Performance object allows setting or returning parameters that influence the performance on multicore processors.
    """
    def __init__(self, app_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.Performance)
        except Exception as e:
            self.__log.error(f'😡 Error initializing Performance object: {str(e)}')

    @property
    def max_num_meas_setup_threads(self):
        """The maximum number of threads CANoe will use.
        """
        return self.com_obj.MaxNumMeasurementSetupThreads

    @max_num_meas_setup_threads.setter
    def max_num_meas_setup_threads(self, num: int):
        """Sets the maximum number of additional threads which may be used for logging branches in the Measurement Setup.
        The property is not writable while the measurement is running.

        Args:
            num (int): The maximum number of threads CANoe will use. By default, this value is calculated from the number of processor cores.
        """
        self.com_obj.MaxNumMeasurementSetupThreads = num