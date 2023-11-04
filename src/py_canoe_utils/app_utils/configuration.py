# Import Python Libraries here
import logging
import pythoncom
import win32com.client
from time import sleep as wait

logger_inst = logging.getLogger('CANOE_LOG')


class Configuration:
    """The Configuration object represents the active configuration.
    """

    def __init__(self, app_com_obj: object, enable_config_events=False):
        self.log = logger_inst
        self.com_obj = win32com.client.Dispatch(app_com_obj.Configuration)
        if enable_config_events:
            win32com.client.WithEvents(self.com_obj, CanoeConfigurationEvents)

    @property
    def comment(self) -> str:
        """Gets the comment for the configuration.

        Returns:
            str: The comment.
        """
        return self.com_obj.Comment

    @comment.setter
    def comment(self, text: str) -> None:
        """Defines the comment for the configuration.

        Args:
            text (str): The comment.
        """
        self.com_obj.Comment = text
        self.log.info(f'configuration comment set to {text}.')

    @property
    def fdx_enabled(self) -> int:
        """Enables/Disables value of FDX protocol.

        Returns:
            int: The activation state of the FDX protocol. 0: FDX protocol is deactivated. 1: FDX protocol is activated.
        """
        return self.com_obj.FDXEnabled

    @fdx_enabled.setter
    def fdx_enabled(self, enabled: int) -> None:
        """Enables/Disables the FDX protocol.

        Args:
            enabled (int): The activation state of the FDX protocol. 0: deactivate FDX protocol. ≠0: activate FDX protocol.
        """
        self.com_obj.FDXEnabled = enabled
        self.log.info(f'FDX protocol set to {enabled}.')

    @property
    def full_name(self) -> str:
        """gets the complete path of the configuration.

        Returns:
            str: complete path of the configuration.
        """
        return self.com_obj.FullName

    @full_name.setter
    def full_name(self, full_name: str) -> None:
        """sets the complete path of the configuration.

        Args:
            full_name (str): The new complete path of the configuration.
        """
        self.com_obj.FullName = full_name
        self.log.info(f'complete path of the configuration set to {full_name}.')

    @property
    def mode(self) -> int:
        """returns whether the Online mode or the Offline mode is active.

        Returns:
            int: The currently active mode.
        """
        return self.com_obj.Mode

    @mode.setter
    def mode(self, mode: int) -> None:
        """sets the Online mode or the Offline mode to active.

        Args:
            mode (int): The active mode; valid values are: 0-Online mode is activated. 1-Offline mode is activated.
        """
        self.com_obj.Mode = mode
        self.log.info(f'offline/online mode set to {mode}.')

    @property
    def modified(self) -> bool:
        """Returns information on whether the current configuration was modified since the time it was loaded or created, or sets this property.
        This property determines whether the user is prompted to save when another configuration is loaded.

        Returns:
            bool: The current value of the property.
        """
        return self.com_obj.Modified

    @modified.setter
    def modified(self, modified: bool) -> None:
        """sets Modified property to flase/true.

        Args:
            modified (bool): Value to be assigned to the Modified property.
        """
        self.com_obj.Modified = modified
        self.log.info(f'configuration modified property set to {modified}.')

    @property
    def name(self) -> str:
        """Returns the name of the configuration.

        Returns:
            str: The name of the currently loaded configuration.
        """
        return self.com_obj.Name

    @property
    def path(self) -> str:
        """returns the path of the configuration, depending on the actual configuration.

        Returns:
            str: The path of the currently loaded configuration.
        """
        return self.com_obj.Path

    @property
    def read_only(self) -> bool:
        """Indicates whether the configuration is write protected.

        Returns:
            bool: If the object is write protected True is returned; otherwise False is returned.
        """
        return self.com_obj.ReadOnly

    @property
    def saved(self) -> bool:
        """Indicates whether changes to the configuration have already been saved.

        Returns:
            bool: If changes were made to the configuration and they have not been saved yet, False is returned; otherwise True is returned.
        """
        return self.com_obj.Saved

    def compile_and_verify(self):
        """Compiles all CAPL test modules and verifies all XML test modules.
        All test modules in the Simulation Setup and in the Test Setup are taken into consideration.
        """
        self.com_obj.CompileAndVerify()
        self.log.info(f'Compiled all test modules in the Simulation Setup and in the Test Setup.')

    def save(self, path='', prompt_user=False):
        """Saves the configuration.

        Args:
            path (str): The complete file name. If no path is specified, the configuration is saved under its current name. If it is not saved yet, the user will be prompted for a name.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations.
        """
        if not self.saved:
            if path == '':
                self.com_obj.Save()
            else:
                self.com_obj.Save(path, prompt_user)
            self.log.info(f'Saved configuration({path}).')
        else:
            self.log.info('CANoe Configuration already in saved state.')
        return self.saved

    def save_as(self, path: str, major: str, minor: str, prompt_user: bool):
        """Saves the configuration as a different CANoe version

        Args:
            path (str): The complete path.
            major (str): The major version number of the target version, e.g. 10 for CANoe 10.1.
            minor (str): The minor version number of the target version, e.g. 1 for CANoe 10.1
            prompt_user (bool): A boolean value that defines whether the user should interact in error situations.
        """
        self.com_obj.SaveAs(path, major, minor, prompt_user)
        self.log.info(f'Saved configuration as {path}.')

    def get_all_test_setup_environments(self) -> dict:
        test_environments_info = dict()
        test_setup_environments = self.com_obj.TestSetup.TestEnvironments
        for test_env in test_setup_environments:
            test_environments_info[test_env.Name] = test_env
        return test_environments_info
    
    def get_all_test_modules_in_test_environment(self, test_environment_com_obj: object) -> dict:
        test_modules_info = dict()
        test_modules = test_environment_com_obj.Items
        for test_module in test_modules:
            test_modules_info[test_module.Name] = test_module
        return test_modules_info


class CanoeConfigurationEvents:
    """Handler for CANoe Configuration events"""

    @staticmethod
    def OnClose():
        """Occurs when the configuration is closed.
        """
        logger_inst.info('configuration OnClose event triggered.')

    @staticmethod
    def OnSystemVariablesDefinitionChanged():
        """Occurs when system variable definitions are added, changed or removed.
        """
        logger_inst.info('configuration OnSystemVariablesDefinitionChanged event triggered.')


class TestModule:
    TM_STARTED = False
    TM_STOPPED = False
    TM_REPORT_GENERATED = False
    TM_REPORT_PATH = ''

    def __init__(self, test_module_com_obj: object):
        self.tm_com_obj = win32com.client.Dispatch(test_module_com_obj)
        self.wait_for_tm_to_start = lambda: TmDoEventsUntil(lambda: TestModule.TM_STARTED)
        self.wait_for_tm_to_stop = lambda: TmDoEventsUntil(lambda: TestModule.TM_STOPPED)
        win32com.client.DispatchWithEvents(self.tm_com_obj, TestModuleEvents)

    @property
    def name(self) -> str:
        """Returns the name of the object.

        Returns:
            str: The name of the TSTestModule object.
        """
        return self.tm_com_obj.Name

    @property
    def full_name(self) -> str:
        """determines the complete path of the object.

        Returns:
            str: The complete path.
        """
        return self.tm_com_obj.FullName

    @property
    def path(self) -> str:
        """returns the path of the object, depending on the actual object.

        Returns:
            str: The complete path of the CAPL program or of the XML test module executed in the test module.
        """
        return self.tm_com_obj.Path

    @property
    def verdict(self) -> int:
        """Returns the verdict of the test tree element(test module).

        Returns:
            int: The verdict of the test tree element. 0=VerdictNotAvailable, 1=VerdictPassed, 2=VerdictFailed.
        """
        return self.tm_com_obj.Verdict

    def start(self):
        """Starts the test module.
        """
        self.tm_com_obj.Start()

    def pause(self):
        pass

    def resume(self):
        pass

    def stop(self):
        pass

    def reload(self):
        pass

    def set_execution_time(self, days: int, hours: int, minutes: int):
        pass


class TestModuleEvents:
    def __init__(self):
        pass

    def OnStart(self):
        TestModule.TM_REPORT_PATH = ''
        TestModule.TM_REPORT_GENERATED = False
        TestModule.TM_STARTED = True
        TestModule.TM_STOPPED = False
        print(f'test module OnStart event.')

    def OnPause(self):        
        print(f'test module OnPause event.')

    def OnStop(self, reason):        
        TestModule.TM_STARTED = False
        TestModule.TM_STOPPED = True
        print(f'test module OnStop event. reason -> {reason}')

    def OnReportGenerated(self, Success, SourceFullName, GeneratedFullName):
        TestModule.TM_REPORT_PATH = SourceFullName
        TestModule.TM_REPORT_GENERATED = Success
        print(f'test module OnReportGenerated event. {Success} # {SourceFullName} # {GeneratedFullName}')

    def OnVerdictFail(self):
        pass

def TmDoEvents():
    pythoncom.PumpWaitingMessages()
    wait(.1)


def TmDoEventsUntil(condition):
    while not condition():
        TmDoEvents()
