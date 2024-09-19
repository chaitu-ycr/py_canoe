# import external modules here
import logging
import win32com.client


class CanoeSimulationEvents:
    """Handler for CANoe Simulation events"""

    @staticmethod
    def OnIdle(time_high: int, time: int) -> None:
        logging.getLogger('CANOE_LOG').debug(f'👉 Simulation OnIdle event triggered. time_high = {time_high} and time = {time}')

    @staticmethod
    def OnIdleU(time_high: int, time: int) -> None:
        logging.getLogger('CANOE_LOG').debug(f'👉 Simulation OnIdleU event triggered. time_high = {time_high} and time = {time}')


class Simulation:
    """The Simulation object represents CANoe's measurement functions in the Simulation mode.
    With the help of the Simulation object you can control the system time from an external source during the measurement.
    """

    def __init__(self, app_com_obj, enable_sim_events=False):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.Simulation)
            if enable_sim_events:
                win32com.client.WithEvents(self.com_obj, CanoeSimulationEvents)
        except Exception as e:
            self.__log.error(f'😡 Error initializing Simulation object: {str(e)}')

    @property
    def animation(self) -> int:
        return self.com_obj.Animation

    @animation.setter
    def animation(self, value: int) -> None:
        self.com_obj.Animation = value
        self.__log.debug(f'👉 animation factor set to = {value}.')

    @property
    def current_time(self) -> int:
        return self.com_obj.CurrentTime

    @property
    def current_time_high(self) -> int:
        return self.com_obj.CurrentTimeHigh

    @property
    def notification_type(self) -> int:
        return self.com_obj.NotificationType

    @notification_type.setter
    def notification_type(self, value: int) -> None:
        self.com_obj.NotificationType = value
        self.__log.debug(f'👉 notification type set to = {value}.')

    def increment_time(self, ticks: int) -> None:
        self.com_obj.IncrementTime(ticks)
        self.__log.debug(f'👉 Increased the system time to = {ticks} ticks.')

    def increment_time_and_wait(self, ticks: int) -> None:
        self.com_obj.IncrementTimeAndWait(ticks)
        self.__log.debug(f'👉 Increased the system time to = {ticks} ticks.')
