# Import Python Libraries here
import logging
import pythoncom
import win32com.client
from time import sleep as wait

logger_inst = logging.getLogger('CANOE_LOG')


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


class Configuration:
    """The Configuration object represents the active configuration.
    """

    def __init__(self, app_com_obj, enable_config_events=False):
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
    def full_name(self, full_name: str):
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
        return self.com_obj.mode

    @mode.setter
    def mode(self, mode: int) -> None:
        """sets the Online mode or the Offline mode to active.

        Args:
            mode (int): The active mode; valid values are: 0-Online mode is activated. 1-Offline mode is activated.
        """
        self.com_obj.mode = mode
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

    @property
    def simulation_setup(self):
        return SimulationSetup(self.com_obj)

    @property
    def test_setup(self):
        return TestSetup(self.com_obj)

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
        return self.test_setup.test_environments.fetch_all_test_environments()

    def get_all_test_modules_in_test_environments(self) -> list:
        test_modules = list()
        tse = self.get_all_test_setup_environments()
        for te_name, te_inst in tse.items():
            for tm_name, tm_inst in te_inst.get_all_test_modules().items():
                test_modules.append({'name': tm_name, 'object': tm_inst, 'environment': te_name})
        return test_modules


class TestSetup:
    def __init__(self, conf_com_obj):
        self.com_obj = win32com.client.Dispatch(conf_com_obj.TestSetup)

    def save_all(self, prompt_user=False) -> None:
        """Saves all test environments of the test setup. If no storage path has been set, the user is prompted for input.

        Args:
            prompt_user (bool, optional): A boolean value that defines whether the user should interact in error situations (optional). Defaults to False.
        """
        self.com_obj.SaveAll(prompt_user)

    @property
    def test_environments(self):
        return TestEnvironments(self.com_obj)


class TestEnvironments:
    def __init__(self, test_setup_com_obj):
        self.com_obj = win32com.client.Dispatch(test_setup_com_obj.TestEnvironments)

    @property
    def count(self) -> int:
        """The number of test environments contained.

        Returns:
            int: Returns the number of test environments.
        """
        return self.com_obj.Count

    def add(self, name: str) -> object:
        """Adds a new test environment to CANoe's Test Setup.
        The path may be absolute or relative to the current CANoe configuration.

        Args:
            name (str): If a new test environment shall be created, name contains the name of the new test environment.
                        If an existing test environment shall be read from a file, name contains the path of this file.
        
        Returns:
            object: The TestEnvironment object of the new test environment.
        """
        return self.com_obj.Add(name)

    def remove(self, index: int, prompt_user=False) -> None:
        """Removes a test environment from CANoe's Test Setup.
        The index can contain the number or the name of the test environment.
        If a number is given, 1 refers to the first test environment, 2 refers to the second test environment,…


        Args:
            index (int): The index of the object to be removed.
            prompt_user (bool, optional): A boolean value that determines whether the user will be prompted before deleting the test environment. Defaults to False.
        """
        self.com_obj.Remove(index, prompt_user)

    def fetch_all_test_environments(self):
        test_environments = dict()
        for index in range(1, self.count + 1):
            te_com_obj = win32com.client.Dispatch(self.com_obj.Item(index))
            te_inst = TestEnvironment(te_com_obj)
            test_environments[te_inst.name] = te_inst
        return test_environments


class TestEnvironment:
    def __init__(self, test_environment_com_obj):
        self.com_obj = test_environment_com_obj
        self.__test_modules = TestModules(self.com_obj)

    @property
    def enabled(self) -> bool:
        """returns whether the object is in an active/inactive state

        Returns:
            bool: A boolean value whether the TestEnvironment activated or not
        """
        return self.com_obj.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        """Activates/deactivates an object.

        Args:
            value (bool): A boolean value that determines whether the TestEnvironment should become activated or not
        """
        self.com_obj.Enabled = value

    @property
    def full_name(self) -> str:
        """The complete path of the test environment.

        Returns:
            str: The complete path of the test environment.
        """
        return self.com_obj.FullName

    @property
    def name(self) -> str:
        """The name of the test environment.

        Returns:
            str: The name of the test environment
        """
        return self.com_obj.Name

    @property
    def path(self) -> str:
        """The complete path to the test environment.

        Returns:
            str: The complete path to the test environment.
        """
        return self.com_obj.Path

    def execute_all(self) -> None:
        """Starts all test modules within the test environment consecutively.
        This means that initially the first test module is started.
        As soon as it finished (completely) the second test module is started, then the third.
        This method can only be used with test environments that exclusively contain test modules.
        """
        self.com_obj.ExecuteAll()

    def save(self, name: str, prompt_user=False) -> None:
        """Saves the test environment.

        Args:
            name (str): Sets the (new) path for the test environment, if applicable. If no path is specified, the test environment is saved under its current name. If it is not saved yet, the user will be prompted for a name.
            prompt_user (bool, optional): A boolean value that determines whether the user will be prompted on error. Defaults to False.
        """
        self.com_obj.Save(name, prompt_user)

    def save_as(self, name: str, major: int, minor: int, prompt_user=False) -> None:
        """Saves the test environment in older formats.
        If you specify "0, 0" (or any invalid or pre 5.1 version) for the "major, minor" parameters, 
        the test environment will be saved in the file format of the CANoe version you are currently running.

        Args:
            name (str): The path of the file in which the test environment will be saved.
            major (int): Prefix of the version number: e.g. 5 with version 5.1.
            minor (int): Suffix of the version number: e.g. 1 with version 5.1
            prompt_user (bool, optional): A boolean value that determines whether the user will be prompted on error. Defaults to False.
        """
        self.com_obj.SaveAs(name, major, minor, prompt_user)

    def stop_sequence(self) -> None:
        """Stops the consecutive execution of all test modules in the test environment.
        This method can only be used with test environments that exclusively contain test modules.
        It may only be called while a measurement and the sequential execution is running.
        """
        self.com_obj.StopSequence()

    def get_all_test_modules(self):
        return self.__test_modules.fetch_test_modules()


class TestModules:
    """The TestModules object represents the test modules in a test environment or in a test setup folder.
    This object should be preferred to the TestSetupItems object.
    """

    def __init__(self, test_env_com_obj) -> None:
        self.com_obj = test_env_com_obj.TestModules

    @property
    def count(self) -> int:
        """The number of test modules contained.

        Returns:
            int: The number of test modules.
        """
        return self.com_obj.Count

    def add(self, full_name: str) -> object:
        """Adds a test module to a test environment or a test setup folder in CANoe's Test Setup.
        The path can be absolute or relative to the current CANoe configuration.
        This method fails if the path is not valid.

        Args:
            full_name (str): The path of the test module specification. This can be a CAPL program or an XML test description for example.

        Returns:
            object: The TSTestModule object of the new test module
        """
        return self.com_obj.Add(full_name)

    def remove(self, index: int, prompt_user=False) -> None:
        """Removes a test module from a test environment or a test setup folder.
        The index can contain the number or the name of the test environment.
        If a number is given, 1 refers to the first folder, 2 refers to the second test environment etc.

        Args:
            index (int): The index of the object to be removed.
            prompt_user (bool, optional): A boolean value that determines whether the user will be prompted be-fore deleting the folder. Defaults to False.
        """
        self.com_obj.Remove(index, prompt_user)

    def fetch_test_modules(self) -> dict:
        test_modules = dict()
        for index in range(1, self.count + 1):
            tm_com_obj = self.com_obj.Item(index)
            tm_inst = TestModule(tm_com_obj)
            test_modules[tm_inst.name] = tm_inst
        return test_modules


def TmDoEvents():
    pythoncom.PumpWaitingMessages()
    wait(.1)


def TmDoEventsUntil(condition):
    while not condition():
        TmDoEvents()


class TestModuleEvents:
    def __init__(self):
        self.tm_html_report_path = ''
        self.tm_report_generated = False
        self.tm_started = False
        self.tm_stopped = False
        self.tm_running = self.tm_started

    def OnStart(self):
        """OnStart is called after the test module started.
        """
        self.tm_html_report_path = ''
        self.tm_report_generated = False
        self.tm_started = True
        self.tm_stopped = False
        self.tm_running = self.tm_started
        # logger_inst.info(f'test module OnStart event.')

    @staticmethod
    def OnPause():
        """OnPause is called when the test module execution has been aborted.
        """
        logger_inst.info(f'test module OnPause event.')

    def OnStop(self, reason):
        """OnStop is called after the execution of a test module is stopped.

        Args:
            reason (int): Contains the cause of the test module execution stop as a value of type TestModuleTestReason.
                            0: The test module was executed completely.
                            1: The test module was stopped by the user.
                            2: The test module was stopped by measurement stop.
        """
        self.tm_started = False
        self.tm_stopped = True
        self.tm_running = self.tm_started
        # logger_inst.info(f'test module OnStop event. reason -> {reason}')

    def OnReportGenerated(self, success, source_full_name, generated_full_name):
        """OnReportGenerated is called after an HTML test report has been generated (successfully or not) from an XML test report.

        Args:
            success (bool): Contains the value true if the HTML test report is generated successfully.
                            If an error occurs, e.g. because the XML test report was not found, this parameter contains the value false
            source_full_name (str): Contains the absolute path for the XML test report from which the HTML test report is to be generated.
                            May be empty in case of an error.
            generated_full_name (str): Contains the absolute path for the generated HTML test report. May be empty in case of an error
        """
        self.tm_html_report_path = generated_full_name
        self.tm_report_generated = success
        logger_inst.info(f'test module OnReportGenerated event. {success} # {source_full_name} # {generated_full_name}')

    @staticmethod
    def OnVerdictFail():
        """OnVerdictFail occurs whenever a test case fails.
        """
        # logger_inst.info(f'test module OnVerdictFail event.')
        pass


class TestModule:

    def __init__(self, test_module_com_obj):
        self.com_obj = win32com.client.DispatchWithEvents(test_module_com_obj, TestModuleEvents)
        self.wait_for_tm_to_start = lambda: TmDoEventsUntil(lambda: self.com_obj.tm_started)
        self.wait_for_tm_to_stop = lambda: TmDoEventsUntil(lambda: self.com_obj.tm_stopped)

    @property
    def name(self) -> str:
        """Returns the name of the object.

        Returns:
            str: The name of the TSTestModule object.
        """
        return self.com_obj.Name

    @property
    def full_name(self) -> str:
        """determines the complete path of the object.

        Returns:
            str: The complete path.
        """
        return self.com_obj.FullName

    @property
    def path(self) -> str:
        """returns the path of the object, depending on the actual object.

        Returns:
            str: The complete path of the CAPL program or of the XML test module executed in the test module.
        """
        return self.com_obj.Path

    @property
    def verdict(self) -> int:
        """Returns the verdict of the test tree element(test module).
        0=VerdictNotAvailable
        1=VerdictPassed
        2=VerdictFailed
        3=VerdictNone (not available for test modules)
        4=VerdictInconclusive (not available for test modules)
        5=VerdictErrorInTestSystem (not available for test modules)

        Returns:
            int: The verdict of the test tree element. 0=VerdictNotAvailable, 1=VerdictPassed, 2=VerdictFailed.
        """
        return self.com_obj.Verdict

    def start(self):
        """Starts the test module.
        """
        self.com_obj.Start()
        logger_inst.info(f'started executing test module. waiting for completion...')
        self.wait_for_tm_to_stop()
        logger_inst.info(f'completed executing test module. verdict = {self.verdict}')

    def pause(self) -> None:
        """Instructs the test module to pause.
        The test module pauses when the test function/control function/test case that has just been executed is terminated.
        The execution dialog can be used to specify whether the test module pauses after test cases or test functions/control functions.
        The OnPause event indicates the actual point at which the test module pauses.
        The execution can be resumed from the point at which the test module paused by means of the Resume method.
        """
        self.com_obj.Pause()

    def resume(self) -> None:
        """Resumes the execution of a suspended test configuration.
        """
        self.com_obj.Resume()

    def stop(self) -> None:
        """Stops the execution of the test module.
        Stopping a test module is only possible at certain points of a test sequence.
        Therefore, as recently as the OnStop event is received one can be sure that the test module really stopped.
        """
        self.com_obj.Stop()
        logger_inst.info(f'stopping test module. waiting for completion...')
        self.wait_for_tm_to_stop()
        logger_inst.info(f'completed stopping test module.')

    def reload(self) -> None:
        """This reloads the XML file with the test specification for XML test modules.
        This method can also be called during a measurement in order to load a modified XML specification.
        The call fails if the affected test module is being executed at that time.
        This method is not implemented for CAPL test modules.
        """
        self.com_obj.Reload()

    def set_execution_time(self, days: int, hours: int, minutes: int):
        pass


class SimulationSetup:
    def __init__(self, conf_com_obj):
        self.com_obj = win32com.client.Dispatch(conf_com_obj.SimulationSetup)

    @property
    def replay_collection(self):
        return ReplayCollection(self.com_obj)

    @property
    def buses(self):
        return Buses(self.com_obj)

    @property
    def nodes(self):
        return Nodes(self.com_obj)


class ReplayCollection:
    """The ReplayCollection object represents the Replay Blocks of the CANoe application.
    """
    def __init__(self, sim_setup_com_obj):
        self.com_obj = win32com.client.Dispatch(sim_setup_com_obj.ReplayCollection)

    @property
    def count(self) -> int:
        """The number of Replay Blocks contained.

        Returns:
            int: The number of Replay Blocks contained.
        """
        return self.com_obj.Count

    def add(self, name: str) -> object:
        """TODO: documentation update pending."""
        return self.com_obj.Add(name)

    def remove(self, index: int) -> None:
        """TODO: documentation update pending."""
        self.com_obj.Remove(index)

    def fetch_replay_blocks(self) -> dict:
        replay_blocks = dict()
        for index in range(1, self.count + 1):
            rb_com_obj = self.com_obj.Item(index)
            rb_inst = ReplayBlock(rb_com_obj)
            replay_blocks[rb_inst.name] = rb_inst
        return replay_blocks


class ReplayBlock:
    def __init__(self, replay_block_com_obj):
        self.com_obj = win32com.client.Dispatch(replay_block_com_obj)

    @property
    def name(self) -> str:
        """The name of the Replay Block."""
        return self.com_obj.Name

    @property
    def path(self) -> str:
        """The path of the replay file."""
        return self.com_obj.Path

    @path.setter
    def path(self, path: str):
        """The path of the replay file."""
        self.com_obj.Path = path

    def start(self):
        """Starts the replay.
        TODO: documentation update pending.
        """
        self.com_obj.Start()

    def stop(self):
        """Stops the replay.
        TODO: documentation update pending.
        """
        self.com_obj.Stop()


class Buses:
    """The Buses object represents the buses of the Simulation Setup of the CANoe application.
    The Buses object is only available in CANoe.
    """
    def __init__(self, sim_setup_com_obj):
        self.com_obj = win32com.client.Dispatch(sim_setup_com_obj.Buses)

    @property
    def count(self) -> int:
        """TODO: documentation update pending."""
        return self.com_obj.Count


class Nodes:
    """The Nodes object represents the CAPL node of the Simulation Setup of the CANoe application.
    The Nodes object is only available in CANoe.
    """
    def __init__(self, sim_setup_com_obj):
        self.com_obj = win32com.client.Dispatch(sim_setup_com_obj.Nodes)

    @property
    def count(self) -> int:
        """TODO: documentation update pending."""
        return self.com_obj.Count
