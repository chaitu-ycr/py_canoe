# import external modules here
import logging
import pythoncom
import win32com.client
from time import sleep as wait

# import internal modules here

class CanoeConfigurationEvents:
    """Handler for CANoe Configuration events"""

    @staticmethod
    def OnClose():
        """Occurs when the configuration is closed.
        """
        logging.getLogger('CANOE_LOG').debug('👉 configuration OnClose event triggered.')

    @staticmethod
    def OnSystemVariablesDefinitionChanged():
        """Occurs when system variable definitions are added, changed or removed.
        """
        logging.getLogger('CANOE_LOG').debug('👉 configuration OnSystemVariablesDefinitionChanged event triggered.')


class Configuration:
    def __init__(self, app_com_obj, enable_config_events=False):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.Configuration)
            if enable_config_events:
                win32com.client.WithEvents(self.com_obj, CanoeConfigurationEvents)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe configuration: {str(e)}')

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
        self.__log.debug(f'👉configuration comment set to {text}.')

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
        self.__log.debug(f'👉FDX protocol set to {enabled}.')

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
        self.__log.debug(f'👉complete path of the configuration set to {full_name}.')

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
        self.__log.debug(f'👉offline/online mode set to {mode}.')

    @property
    def modified(self) -> bool:
        """returns information on whether the current configuration was modified since the time it was loaded or created.

        Returns:
            bool: True if the configuration has been changed, False otherwise.
        """
        return self.com_obj.Modified

    @modified.setter
    def modified(self, value: bool):
        """sets the modified state of the configuration.

        Args:
            value (bool): False to discard any active modification, True otherwise.
        """
        self.com_obj.Modified = value
        self.__log.debug(f"Configuration modified property value set to {value}.")

    @property
    def name(self) -> str:
        """Returns the name of the configuration.

        Returns:
            str: The name of the currently loaded configuration.
        """
        return self.com_obj.Name
    
    @property
    def offline_setup(self):
        return OfflineSetup(self.com_obj)

    @property
    def online_setup(self):
        return OnlineSetup(self.com_obj)
    
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
            bool: False is returned, If changes were made to the configuration and not been saved yet. otherwise True is returned.
        """
        return self.com_obj.Saved
    
    @property
    def simulation_setup(self):
        return SimulationSetup(self.com_obj)

    @property
    def test_setup(self):
        """Returns the TestSetup object.
        """
        return TestSetup(self.com_obj)

    def compile_and_verify(self):
        """Compiles all CAPL test modules and verifies all XML test modules.
        All test modules in the Simulation Setup and in the Test Setup are taken into consideration.
        """
        self.com_obj.CompileAndVerify()
        self.__log.debug(f'👉Compiled all test modules in the Simulation Setup and in the Test Setup.')

    def get_all_test_setup_environments(self) -> dict:
        """returns all test setup environments.
        """
        return self.test_setup.test_environments.fetch_all_test_environments()

    def get_all_test_modules_in_test_environments(self) -> list:
        """returns all test setup modules.
        """
        test_modules = list()
        tse = self.get_all_test_setup_environments()
        for te_name, te_inst in tse.items():
            for tm_name, tm_inst in te_inst.get_all_test_modules().items():
                test_modules.append({'name': tm_name, 'object': tm_inst, 'environment': te_name})
        return test_modules

    def save(self, path='', prompt_user=False):
        """Saves the configuration.

        Args:
            path (str): The complete file name. If no path is specified, the configuration is saved under its current name. If it is not saved yet, the user will be prompted for a name.
            prompt_user (bool, optional): A boolean value that indicates whether the user should intervene in error situations.
        """
        if path == '':
            self.com_obj.Save()
        else:
            self.com_obj.Save(path, prompt_user)
            self.__log.debug(f'👉Saved configuration({path}).')
        return self.saved

    def save_as(self, path: str, major: int, minor: int, prompt_user: bool):
        """Saves the configuration as a different CANoe version

        Args:
            path (str): The complete path.
            major (int): The major version number of the target version, e.g. 10 for CANoe 10.1.
            minor (int): The minor version number of the target version, e.g. 1 for CANoe 10.1
            prompt_user (bool): A boolean value that defines whether the user should interact in error situations.
        """
        self.com_obj.SaveAs(path, major, minor, prompt_user)
        self.__log.debug(f'👉Saved configuration as {path}.')
        return self.saved


class TestSetup:
    """The TestSetup object represents CANoe's test setup.
    """
    def __init__(self, conf_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(conf_com_obj.TestSetup)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe test setup: {str(e)}')

    def save_all(self, prompt_user=False) -> None:
        """Saves all test environments of the test setup. If no storage path has been set, the user is prompted for input.

        Args:
            prompt_user (bool, optional): A boolean value that defines whether the user should interact in error situations (optional). Defaults to False.
        """
        self.com_obj.SaveAll(prompt_user)

    @property
    def test_environments(self):
        """Returns the TestEnvironments object.
        """
        return TestEnvironments(self.com_obj)


class TestEnvironments:
    def __init__(self, test_setup_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(test_setup_com_obj.TestEnvironments)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe test environments: {str(e)}')

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

    def fetch_all_test_environments(self) -> dict:
        """returns all test setup test environments.
        """
        test_environments = dict()
        for index in range(1, self.count + 1):
            te_com_obj = win32com.client.Dispatch(self.com_obj.Item(index))
            te_inst = TestEnvironment(te_com_obj)
            test_environments[te_inst.name] = te_inst
        return test_environments


class TestEnvironment:
    """The TestEnvironment object represents a test environment within CANoe's test setup.
    """
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
            name (str): Sets the (new) path for the test environment, if applicable.
                If no path is specified, the test environment is saved under its current name.
                If it is not saved yet, the user will be prompted for a name.
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
        """returns all test modules in a test environment.
        """
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
        """returns all test modules in a test environment.
        """
        test_modules = dict()
        for index in range(1, self.count + 1):
            tm_com_obj = self.com_obj.Item(index)
            tm_inst = TestModule(tm_com_obj)
            test_modules[tm_inst.name] = tm_inst
        return test_modules


def TmDoEvents():
    """pumps wait message and waits for 100 ms."""
    pythoncom.PumpWaitingMessages()
    wait(.1)


def TmDoEventsUntil(condition):
    """triggers wait event every 100ms till condition is satisfied."""
    while not condition():
        TmDoEvents()


class TestModuleEvents:
    """test module events object.
    """
    def __init__(self):
        self.tm_html_report_path = ''
        self.tm_report_generated = False
        self.tm_running = False

    def OnStart(self):
        """OnStart is called after the test module started.
        """
        self.tm_html_report_path = ''
        self.tm_report_generated = False
        self.tm_running = True
        # logging.getLogger('CANOE_LOG').debug(f'👉test module OnStart event.')

    @staticmethod
    def OnPause():
        """OnPause is called when the test module execution has been aborted.
        """
        logging.getLogger('CANOE_LOG').debug(f'👉test module OnPause event.')

    def OnStop(self, reason):
        """OnStop is called after the execution of a test module is stopped.

        Args:
            reason (int): Contains the cause of the test module execution stop as a value of type TestModuleTestReason.
                            0: The test module was executed completely.
                            1: The test module was stopped by the user.
                            2: The test module was stopped by measurement stop.
        """
        self.tm_running = False
        # logging.getLogger('CANOE_LOG').debug(f'👉test module OnStop event. reason -> {reason}')

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
        self.tm_running = False
        logging.getLogger('CANOE_LOG').debug(f'👉test module OnReportGenerated event. {success} # {source_full_name} # {generated_full_name}')

    def OnVerdictFail(self):
        """OnVerdictFail occurs whenever a test case fails.
        """
        # logging.getLogger('CANOE_LOG').debug(f'👉test module OnVerdictFail event.')
        pass


class TestModule:
    """The TestModule object represents a test module in CANoe's test setup.
    """

    def __init__(self, test_module_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.DispatchWithEvents(test_module_com_obj, TestModuleEvents)
            self.wait_for_tm_to_start = lambda: TmDoEventsUntil(lambda: self.com_obj.tm_running)
            self.wait_for_tm_to_stop = lambda: TmDoEventsUntil(lambda: not self.com_obj.tm_running)
            self.wait_for_tm_report_gen = lambda: TmDoEventsUntil(lambda: self.com_obj.tm_report_generated)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe test module: {str(e)}')

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
        self.wait_for_tm_to_start()
        logging.getLogger('CANOE_LOG').debug(f'👉 started executing test module. waiting for completion...')
    
    def wait_for_completion(self):
        """waits for test module execution completion."""
        self.wait_for_tm_to_stop()
        wait(1)
        logging.getLogger('CANOE_LOG').debug(f'👉 completed executing test module. verdict = {self.verdict}')

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
        logging.getLogger('CANOE_LOG').debug(f'👉stopping test module. waiting for completion...')
        self.wait_for_tm_to_stop()
        logging.getLogger('CANOE_LOG').debug(f'👉completed stopping test module.')

    def reload(self) -> None:
        """This reloads the XML file with the test specification for XML test modules.
        This method can also be called during a measurement in order to load a modified XML specification.
        The call fails if the affected test module is being executed at that time.
        This method is not implemented for CAPL test modules.
        """
        self.com_obj.Reload()

    def set_execution_time(self, days: int, hours: int, minutes: int):
        """Sets the amount of time to run the test configuration repeatedly.

        Args:
            days (int): Specifies for how many days the test module is executed.
            hours (int): Specifies for how many hours the test module is executed (additionally).
            minutes (int): Specifies for how many minutes the test module is executed (additionally).
        """
        self.com_obj.SetExecutionTime(days, hours, minutes)


class OfflineSetup:
    def __init__(self, conf_com_obj) -> None:
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(conf_com_obj.OfflineSetup)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe offline setup: {str(e)}')

    @property
    def source(self) -> object:
        """The source object of the offline setup.

        Returns:
            object: The source object of the offline setup.
        """
        return Source(self.com_obj)


class Source:
    def __init__(self, offlince_setup_com_obj) -> None:
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(offlince_setup_com_obj.Source)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe offline setup source: {str(e)}')

    @property
    def full_name(self) -> str:
        """Returns the complete path of the source.

        Returns:
            str: The complete path of the source.
        """
        return self.com_obj.FullName

    @property
    def sources(self) -> object:
        """Returns A Files object.
        """
        return Sources(self.com_obj)

    @property
    def time_section(self) -> object:
        """Returns A TimeSection object.
        """
        return TimeSection(self.com_obj)

    def export_mapping_table(self, file_name: str) -> None:
        """Exports the current channel mapping configuration to an XML file with the given file name.

        Args:
            file_name (str): The file name to export the mapping table.
        """
        self.com_obj.ExportMappingTable(file_name)
        self.__log.debug(f"Exported mapping table to {file_name}.")

    def get_mapping_table(self, type: int) -> object:
        """Returns the mapping table as a string.

        Args:
            type (int): The type of mapping table to return.

        Returns:
            object: The mapping table object.
        """
        return self.com_obj.GetMappingTable(type)

    def get_mapping_table_by_name(self, type_name: str) -> object:
        """Returns the mapping table as a string.

        Args:
            type_name (str): The name of the mapping table to return.

        Returns:
            object: The mapping table object.
        """
        return self.com_obj.GetMappingTableByName(type_name)

    def import_mapping_table(self, file_name: str) -> None:
        """Imports the channel mapping configuration from an XML file with the given file name.

        Args:
            file_name (str): The file name to import the mapping table.
        """
        self.com_obj.ImportMappingTable(file_name)
        self.__log.debug(f"Imported mapping table from {file_name}.")


class Sources:
    def __init__(self, source_com_obj) -> None:
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(source_com_obj.Sources)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe offline setup source: {str(e)}')

    @property
    def count(self) -> int:
        """Returns the number of sources.

        Returns:
            int: The number of sources.
        """
        return self.com_obj.Count

    @property
    def paths(self) -> list:
        """Returns the paths of the sources.

        Returns:
            list: The paths of the sources.
        """
        list_of_paths = []
        for index in range(1, self.count + 1):
            list_of_paths.append(self.com_obj.Item(index))
        return list_of_paths

    def add(self, source_file: str) -> object:
        """Adds a source file to the offline setup.

        Args:
            source_file (str): The source file to be added.

        Returns:
            object: The added source file.
        """
        return self.com_obj.Add(source_file)

    def clear(self) -> None:
        """Removes all files from the collection."""
        self.com_obj.Clear()

    def remove(self, index: int) -> None:
        """Removes a source file from the offline setup.

        Args:
            index (index): The index of source file to be removed.
        """
        self.com_obj.Remove(index)


class TimeSection:
    """The TimeSection object represents the time section that will be considered while replaying a file in offline mode or converting a file using the applications logging converter features."""
    def __init__(self, source_com_obj) -> None:
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(source_com_obj.TimeSection)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe offline setup source: {str(e)}')

    @property
    def end(self) -> str:
        """Returns the end time of the time section.

        Returns:
            str: The end time of the time section.
        """
        return self.com_obj.End

    @end.setter
    def end(self, time: str) -> None:
        """Sets the end time of the time section.

        Args:
            time (str): The end time of the time section.
        """
        self.com_obj.End = time
        self.__log.debug(f"Time section end time set to {time}.")

    @property
    def start(self) -> str:
        """Returns the start time of the time section.

        Returns:
            str: The start time of the time section.
        """
        return self.com_obj.Start

    @start.setter
    def start(self, time: str) -> None:
        """Sets the start time of the time section.

        Args:
            time (str): The start time of the time section.
        """
        self.com_obj.Start = time
        self.__log.debug(f"Time section start time set to {time}.")

    @property
    def type(self) -> int:
        """Returns the type of the time section.

        Returns:
            int: The type of the time section.
        """
        return self.com_obj.Type


class OnlineSetup:
    def __init__(self, conf_com_obj) -> None:
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(conf_com_obj.OnlineSetup)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe online setup: {str(e)}')

    @property
    def bus_statistics(self) -> object:
        """Returns the BusStatistics object.

        Returns:
            object: The BusStatistics object.
        """
        return BusStatistics(self.com_obj)


class BusStatistics:
    """The BusStatistics object represents the bus statistics of the CANoe application."""
    def __init__(self, setup_com_obj) -> None:
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(setup_com_obj.BusStatistics)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe bus statistics: {str(e)}')

    def bus_statistic(self, bus_type: int, channel: int) -> object:
        """Returns a CANBusStatistic object.

        Args:
            bus_type (int): The bus type.
            channel (int): The channel number.

        Returns:
            object: A CANBusStatistic object.
        """
        return BusStatistic(self.com_obj, bus_type, channel)


class BusStatistic:
    """Returns a CANBusStatistic object."""
    def __init__(self, bus_statistics_com_obj, bus_type: int, channel: int) -> None:
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(bus_statistics_com_obj.BusStatistic(bus_type, channel))
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe bus statistic: {str(e)}')

    @property
    def bus_load(self):
        return self.com_obj.BusLoad

    @property
    def chip_state(self):
        return self.com_obj.ChipState

    @property
    def error(self):
        return self.com_obj.Error

    @property
    def error_total(self):
        return self.com_obj.ErrorTotal

    @property
    def extended(self):
        return self.com_obj.Extended

    @property
    def extended_remote(self):
        return self.com_obj.ExtendedRemote

    @property
    def extended_remote_total(self):
        return self.com_obj.ExtendedRemoteTotal

    @property
    def extended_total(self):
        return self.com_obj.ExtendedTotal

    @property
    def overload(self):
        return self.com_obj.Overload

    @property
    def overload_total(self):
        return self.com_obj.OverloadTotal

    @property
    def peak_load(self):
        return self.com_obj.PeakLoad

    @property
    def rx_error_count(self):
        return self.com_obj.RxErrorCount

    @property
    def standard(self):
        return self.com_obj.Standard

    @property
    def standard_remote(self):
        return self.com_obj.StandardRemote

    @property
    def standard_remote_total(self):
        return self.com_obj.StandardRemoteTotal

    @property
    def standard_total(self):
        return self.com_obj.StandardTotal

    @property
    def tx_error_count(self):
        return self.com_obj.TxErrorCount


class SimulationSetup:
    """The SimulationSetup object represents the Simulation Setup of CANoe.
    """
    def __init__(self, conf_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(conf_com_obj.SimulationSetup)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe simulation setup: {str(e)}')

    @property
    def replay_collection(self):
        """Returns the ReplayCollection object.
        """
        return ReplayCollection(self.com_obj)

    @property
    def buses(self):
        """The Buses object represents the buses of the Simulation Setup of the CANoe application.
        The Buses object is only available in CANoe.
        """
        return Buses(self.com_obj)

    @property
    def nodes(self):
        """Returns the Nodes object.
        """
        return Nodes(self.com_obj)


class ReplayCollection:
    """The ReplayCollection object represents the Replay Blocks of the CANoe application.
    """
    def __init__(self, sim_setup_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(sim_setup_com_obj.ReplayCollection)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe replay collection: {str(e)}')

    @property
    def count(self) -> int:
        """The number of Replay Blocks contained.

        Returns:
            int: The number of Replay Blocks contained.
        """
        return self.com_obj.Count

    def add(self, name: str) -> object:
        """adds a new replay block.

        Args:
            name (str): name of new replay block.

        Returns:
            object: replay block com object.
        """
        return self.com_obj.Add(name)

    def remove(self, index: int) -> None:
        """remove replay block by index.

        Args:
            index (int): index value of replay block to be removed.
        """
        self.com_obj.Remove(index)

    def fetch_replay_blocks(self) -> dict:
        """returns all replay blocks in configuration.
        """
        replay_blocks = dict()
        for index in range(1, self.count + 1):
            rb_com_obj = self.com_obj.Item(index)
            rb_inst = ReplayBlock(rb_com_obj)
            replay_blocks[rb_inst.name] = rb_inst
        return replay_blocks


class ReplayBlock:
    def __init__(self, replay_block_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(replay_block_com_obj)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe replay block: {str(e)}')

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
        """The path of the replay file.

        Args:
            path (str): new path off replay block.
        """
        self.com_obj.Path = path

    def start(self):
        """Starts the replay block.
        """
        self.com_obj.Start()

    def stop(self):
        """Stops the replay block.
        """
        self.com_obj.Stop()


class Buses:
    """The Buses object represents the buses of the Simulation Setup of the CANoe application.
    The Buses object is only available in CANoe.
    """
    def __init__(self, sim_setup_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(sim_setup_com_obj.Buses)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe buses: {str(e)}')

    @property
    def count(self) -> int:
        """returns the number of buses contained.
        """
        return self.com_obj.Count


class Nodes:
    """The Nodes object represents the CAPL node of the Simulation Setup of the CANoe application.
    The Nodes object is only available in CANoe.
    """
    def __init__(self, sim_setup_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(sim_setup_com_obj.Nodes)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe nodes: {str(e)}')

    @property
    def count(self) -> int:
        """returns the number of nodes contained.
        """
        return self.com_obj.Count
