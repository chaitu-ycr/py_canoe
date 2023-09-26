# Import Python Libraries here
import win32com.client

class Simulation:
    """The Simulation object represents CANoe's measurement functions in the Simulation mode.
    With the help of the Simulation object you can control the system time from an external source during the measurement.
    """
    def __init__(self, app) -> None:
        self.app = app
        self.log = self.app.log
        self.sim_obj = win32com.client.Dispatch(self.app.app_com_obj.Simulation)
        win32com.client.WithEvents(self.meas_obj, CanoeSimulationEvents)
    
    @property
    def animation(self) -> int:
        """Returns the animation factor.

        Returns:
            int: The animation factor.
        """
        return self.sim_obj.Animation
    
    @animation.setter
    def animation(self, value: int) -> None:
        """Sets the animation factor.

        Args:
            value (int): The animation factor.
        """
        self.sim_obj.Animation = value
        self.log.info(f'animation factor set to = {value}.')

    @property
    def current_time(self) -> int:
        """Returns the low-order 32 bit of the current system time.

        Returns:
            int: The low-order 32 bit of the current system time.
        """
        return self.sim_obj.CurrentTime

    @property
    def current_time_high(self) -> int:
        """Returns the high-order 32 bit of the current system time.

        Returns:
            int: The high-order 32 bit of the current system time.
        """
        return self.sim_obj.CurrentTimeHigh

    @property
    def notification_type(self) -> int:
        """gets the notification type for the OnIdle handler.

        Returns:
            int: The notification type. 0-Idle. 1-IdleU
        """
        return self.sim_obj.NotificationType
    
    @notification_type.setter
    def notification_type(self, value: int) -> None:
        """sets the notification type for the OnIdle handler.

        Args:
            value (int): The notification type. 0-Idle. 1-IdleU
        """
        self.sim_obj.NotificationType = value
        self.log.info(f'notification type set to = {value}.')

    def increment_time(self, ticks: int) -> None:
        """Increases the system time during the measurement in simulation mode.
        Whereas IncrementTime returns immediately after calling, the IncrementTimeAndWait function returns not until the simulation step is done in CANoe.

        Args:
            ticks (int): The number of ticks by which the system time is to be increased. 1 tick corresponds to 10 µs.
        """
        self.sim_obj.IncrementTime(ticks)
        self.log.info(f'Increased the system time to = {ticks} ticks.')

    def increment_time_and_wait(self, ticks: int) -> None:
        """Increases the system time during the measurement in Simulation mode.
        Whereas IncrementTime returns immediately after calling, the IncrementTimeAndWait method returns not until the simulation step is done in CANoe.

        Args:
            ticks (int): The number of ticks by which the system time is to be increased. 1 tick corresponds to 10 µs.
        """
        self.sim_obj.IncrementTimeAndWait(ticks)
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
        print(f'Simulation OnIdle event triggered. time_high = {time_high} and time = {time}')
    
    @staticmethod
    def OnIdleU(time_high: int, time: int) -> None:
        """Occurs after a simulation step.

        Args:
            time_high (int): High-order 32 bit of the current simulation time.
            time (int): Low-order 32 bit of the current simulation time.
        """
        print(f'Simulation OnIdleU event triggered. time_high = {time_high} and time = {time}')
