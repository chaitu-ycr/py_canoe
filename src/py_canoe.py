# Standard libraries
import os
import pythoncom
import win32com.client
from time import sleep as wait


class CanoeApp:
    """Class for controlling CANoe via python COM

    Examples:
        >>> # Import CANoe application
        >>> from py_canoe import CanoeApp
        >>> canoe_app = CanoeApp()
    """
    # class variables
    CANOE_APP_NAME = 'CANoe.Application'
    CANOE_DELAY = 2

    def __init__(self):
        self.canoe_cfg_name = None
        self.canoe_obj = None

    def open_cfg(self, cfg_name: str, visible=True) -> None:
        """Method to open CANoe configuration.

        Args:
            cfg_name: CANoe configuration file name including path
            visible (bool): True if you want to see CANoe UI default is True

        Examples:
            >>> canoe_app.open_cfg(cfg_name='demo_canoe.cfg')
        """
        if not os.path.isfile(cfg_name):
            print(f'CANoe cfg "{cfg_name}" not found.')
        print(f'{cfg_name} found on machine.')
        pythoncom.CoInitialize()
        self.canoe_obj = win32com.client.Dispatch(self.CANOE_APP_NAME)
        self.canoe_obj.Visible = visible
        self.canoe_obj.Open(cfg_name)
        self.canoe_cfg_name = cfg_name
        print('CANoe cfg opened and ready to use')

    def close(self, save_cfg_before_close=True) -> None:
        """Method for closing CANoe application

        Args:
            save_cfg_before_close (bool): saves CANoe configuration if True. Default value True.
            
        Examples:
            >>> canoe_app.close()
        """
        if self.check_simulation_running():
            self.stop_simulation()
        if save_cfg_before_close:
            self.save_configuration()
            wait(self.CANOE_DELAY)
        self.canoe_obj.Quit()
        print('CANoe Closed')

    def save_configuration(self) -> None:
        """Method for saving CANoe configuration

        Examples:
            >>> canoe_app.save_configuration()
        """
        if not self.canoe_obj.Configuration.Saved:
            self.canoe_obj.Configuration.Save()
            print('CANoe Cfg saved')

    def start_simulation(self) -> None:
        """Method for starting CANoe simulation

        Examples:
            >>> canoe_app.start_simulation()
        """
        if not self.check_simulation_running():
            self.canoe_obj.Measurement.Start()
            wait(self.CANOE_DELAY)
            for i in range(10):
                if not self.canoe_obj.Measurement.Running:
                    print('waiting for CANoe simulation to start')
                    wait(self.CANOE_DELAY)
                else:
                    break
            print('CANoe simulation started')
        else:
            print('CANoe simulation running')

    def stop_simulation(self) -> None:
        """Method for stopping CANoe simulation

        Examples:
            >>> canoe_app.stop_simulation()
        """
        if self.check_simulation_running():
            self.canoe_obj.Measurement.Stop()
            wait(self.CANOE_DELAY)
            for i in range(10):
                if self.canoe_obj.Measurement.Running:
                    print('CANoe Simulation still running')
                    wait(self.CANOE_DELAY)
                else:
                    break
        print('CANoe simulation stopped')

    def check_simulation_running(self) -> bool:
        """Method for checking CANoe simulation running

        Returns:
            simulation status. `True` if simulation running.

        Examples:
            >>> canoe_app.check_simulation_running()
        """
        return self.canoe_obj.Measurement.Running

    def set_replay_block_file(self, block_name: str, recording_file_path: str) -> None:
        """
        Method for setting CANoe replay block file.

        Args:
            block_name: CANoe replay block name
            recording_file_path: CANoe replay recording file including path.

        Examples:
            >>> canoe_app.set_replay_block_file(block_name='replay block name', recording_file_path='replay file includding path')
        """
        if self.check_simulation_running():
            self.stop_simulation()
        count = self.canoe_obj.Bus.ReplayCollection.Count
        try:
            for i in range(1, count + 1):
                name = self.canoe_obj.Bus.ReplayCollection.Item(i).Name
                if name == block_name:
                    self.canoe_obj.Bus.ReplayCollection.Item(i).Path = recording_file_path
                    print(f'Replay block "{block_name}" updated with "{recording_file_path}" path.')
                    self.save_configuration()
        except Exception as msg:
            print(f'Exception "{msg}" received when setting replay block file.')
