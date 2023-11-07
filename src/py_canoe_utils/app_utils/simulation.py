# Import Python Libraries here
import logging
import win32com.client

logger_inst = logging.getLogger('CANOE_LOG')


class Simulation:
    """The Simulation object represents CANoe's measurement functions in the Simulation mode.
    With the help of the Simulation object you can control the system time from an external source during the measurement.
    """

    def __init__(self, app_com_obj, enable_sim_events=False):
        self.log = logger_inst
        self.com_obj = win32com.client.Dispatch(app_com_obj.Simulation)
        if enable_sim_events:
            win32com.client.WithEvents(self.com_obj, CanoeSimulationEvents)

    @property
    def animation(self) -> int:
        """Returns the animation factor.

        Returns:
            int: The animation factor.
        """
        return self.com_obj.Animation

    @animation.setter
    def animation(self, value: int) -> None:
        """Sets the animation factor.

        Args:
            value (int): The animation factor.
        """
        self.com_obj.Animation = value
        self.log.info(f'animation factor set to = {value}.')

    @property
    def current_time(self) -> int:
        """Returns the low-order 32 bit of the current system time.

        Returns:
            int: The low-order 32 bit of the current system time.
        """
        return self.com_obj.CurrentTime

    @property
    def current_time_high(self) -> int:
        """Returns the high-order 32 bit of the current system time.

        Returns:
            int: The high-order 32 bit of the current system time.
        """
        return self.com_obj.CurrentTimeHigh

    @property
    def notification_type(self) -> int:
        """gets the notification type for the OnIdle handler.

        Returns:
            int: The notification type. 0-Idle. 1-IdleU
        """
        return self.com_obj.NotificationType

    @notification_type.setter
    def notification_type(self, value: int) -> None:
        """sets the notification type for the OnIdle handler.

        Args:
            value (int): The notification type. 0-Idle. 1-IdleU
        """
        self.com_obj.NotificationType = value
        self.log.info(f'notification type set to = {value}.')

    def increment_time(self, ticks: int) -> None:
        """Increases the system time during the measurement in simulation mode.
        Whereas IncrementTime returns immediately after calling, the IncrementTimeAndWait function returns not until the simulation step is done in CANoe.

        Args:
            ticks (int): The number of ticks by which the system time is to be increased. 1 tick corresponds to 10 µs.
        """
        self.com_obj.IncrementTime(ticks)
        self.log.info(f'Increased the system time to = {ticks} ticks.')

    def increment_time_and_wait(self, ticks: int) -> None:
        """Increases the system time during the measurement in Simulation mode.
        Whereas IncrementTime returns immediately after calling, the IncrementTimeAndWait method returns not until the simulation step is done in CANoe.

        Args:
            ticks (int): The number of ticks by which the system time is to be increased. 1 tick corresponds to 10 µs.
        """
        self.com_obj.IncrementTimeAndWait(ticks)
        self.log.info(f'Increased the system time to = {ticks} ticks.')


class CanoeSimulationEvents:
    """Handler for CANoe Simulation events"""

    @staticmethod
    def OnIdle(time_high: int, time: int) -> None:
        """Occurs after a simulation step.

        Args:
            time_high (int): High-order 32 bit of the current simulation time.
            time (int): Low-order 32 bit of the current simulation time.
        """
        logger_inst.info(f'Simulation OnIdle event triggered. time_high = {time_high} and time = {time}')

    @staticmethod
    def OnIdleU(time_high: int, time: int) -> None:
        """Occurs after a simulation step.

        Args:
            time_high (int): High-order 32 bit of the current simulation time.
            time (int): Low-order 32 bit of the current simulation time.
        """
        logger_inst.info(f'Simulation OnIdleU event triggered. time_high = {time_high} and time = {time}')
