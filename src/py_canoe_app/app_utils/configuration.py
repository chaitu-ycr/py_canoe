# import external modules here
import logging
import pythoncom
import win32com.client
from time import sleep as wait


class CanoeConfigurationEvents:
    """Handler for CANoe Configuration events"""

    @staticmethod
    def OnClose():
        logging.getLogger('CANOE_LOG').debug('👉 configuration OnClose event triggered.')

    @staticmethod
    def OnSystemVariablesDefinitionChanged():
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
        return self.com_obj.Comment

    @comment.setter
    def comment(self, text: str) -> None:
        self.com_obj.Comment = text
        self.__log.debug(f'👉configuration comment set to {text}.')

    @property
    def fdx_enabled(self) -> int:
        return self.com_obj.FDXEnabled

    @fdx_enabled.setter
    def fdx_enabled(self, enabled: int) -> None:
        self.com_obj.FDXEnabled = enabled
        self.__log.debug(f'👉FDX protocol set to {enabled}.')

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @full_name.setter
    def full_name(self, full_name: str):
        self.com_obj.FullName = full_name
        self.__log.debug(f'👉complete path of the configuration set to {full_name}.')

    @property
    def mode(self) -> int:
        return self.com_obj.mode

    @mode.setter
    def mode(self, mode: int) -> None:
        self.com_obj.mode = mode
        self.__log.debug(f'👉offline/online mode set to {mode}.')

    @property
    def modified(self) -> bool:
        return self.com_obj.Modified

    @modified.setter
    def modified(self, value: bool):
        self.com_obj.Modified = value
        self.__log.debug(f"Configuration modified property value set to {value}.")

    @property
    def name(self) -> str:
        return self.com_obj.Name
    
    @property
    def offline_setup(self):
        return OfflineSetup(self.com_obj)

    @property
    def online_setup(self):
        return OnlineSetup(self.com_obj)
    
    @property
    def path(self) -> str:
        return self.com_obj.Path

    @property
    def read_only(self) -> bool:
        return self.com_obj.ReadOnly

    @property
    def saved(self) -> bool:
        return self.com_obj.Saved
    
    @property
    def simulation_setup(self):
        return SimulationSetup(self.com_obj)

    @property
    def test_setup(self):
        return TestSetup(self.com_obj)

    def compile_and_verify(self):
        self.com_obj.CompileAndVerify()
        self.__log.debug(f'👉Compiled all test modules in the Simulation Setup and in the Test Setup.')

    def get_all_test_setup_environments(self) -> dict:
        return self.test_setup.test_environments.fetch_all_test_environments()

    def get_all_test_modules_in_test_environments(self) -> list:
        test_modules = list()
        tse = self.get_all_test_setup_environments()
        for te_name, te_inst in tse.items():
            for tm_name, tm_inst in te_inst.get_all_test_modules().items():
                test_modules.append({'name': tm_name, 'object': tm_inst, 'environment': te_name})
        return test_modules

    def save(self, path='', prompt_user=False):
        if path == '':
            self.com_obj.Save()
        else:
            self.com_obj.Save(path, prompt_user)
            self.__log.debug(f'👉Saved configuration({path}).')
        return self.saved

    def save_as(self, path: str, major: int, minor: int, prompt_user: bool):
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
        self.com_obj.SaveAll(prompt_user)

    @property
    def test_environments(self):
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
        return self.com_obj.Count

    def add(self, name: str) -> object:
        return self.com_obj.Add(name)

    def remove(self, index: int, prompt_user=False) -> None:
        self.com_obj.Remove(index, prompt_user)

    def fetch_all_test_environments(self) -> dict:
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
        return self.com_obj.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.com_obj.Enabled = value

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def path(self) -> str:
        return self.com_obj.Path

    def execute_all(self) -> None:
        self.com_obj.ExecuteAll()

    def save(self, name: str, prompt_user=False) -> None:
        self.com_obj.Save(name, prompt_user)

    def save_as(self, name: str, major: int, minor: int, prompt_user=False) -> None:
        self.com_obj.SaveAs(name, major, minor, prompt_user)

    def stop_sequence(self) -> None:
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
        return self.com_obj.Count

    def add(self, full_name: str) -> object:
        return self.com_obj.Add(full_name)

    def remove(self, index: int, prompt_user=False) -> None:
        self.com_obj.Remove(index, prompt_user)

    def fetch_test_modules(self) -> dict:
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
        self.tm_html_report_path = ''
        self.tm_report_generated = False
        self.tm_running = True
        # logging.getLogger('CANOE_LOG').debug(f'👉test module OnStart event.')

    @staticmethod
    def OnPause():
        logging.getLogger('CANOE_LOG').debug(f'👉test module OnPause event.')

    def OnStop(self, reason):
        self.tm_running = False
        # logging.getLogger('CANOE_LOG').debug(f'👉test module OnStop event. reason -> {reason}')

    def OnReportGenerated(self, success, source_full_name, generated_full_name):
        self.tm_html_report_path = generated_full_name
        self.tm_report_generated = success
        self.tm_running = False
        logging.getLogger('CANOE_LOG').debug(f'👉test module OnReportGenerated event. {success} # {source_full_name} # {generated_full_name}')

    def OnVerdictFail(self):
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
        return self.com_obj.Name

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @property
    def path(self) -> str:
        return self.com_obj.Path

    @property
    def verdict(self) -> int:
        return self.com_obj.Verdict

    def start(self):
        self.com_obj.Start()
        self.wait_for_tm_to_start()
        logging.getLogger('CANOE_LOG').debug(f'👉 started executing test module. waiting for completion...')
    
    def wait_for_completion(self):
        self.wait_for_tm_to_stop()
        wait(1)
        logging.getLogger('CANOE_LOG').debug(f'👉 completed executing test module. verdict = {self.verdict}')

    def pause(self) -> None:
        self.com_obj.Pause()

    def resume(self) -> None:
        self.com_obj.Resume()

    def stop(self) -> None:
        self.com_obj.Stop()
        logging.getLogger('CANOE_LOG').debug(f'👉stopping test module. waiting for completion...')
        self.wait_for_tm_to_stop()
        logging.getLogger('CANOE_LOG').debug(f'👉completed stopping test module.')

    def reload(self) -> None:
        self.com_obj.Reload()

    def set_execution_time(self, days: int, hours: int, minutes: int):
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
        return self.com_obj.FullName

    @property
    def sources(self) -> object:
        return Sources(self.com_obj)

    @property
    def time_section(self) -> object:
        return TimeSection(self.com_obj)

    def export_mapping_table(self, file_name: str) -> None:
        self.com_obj.ExportMappingTable(file_name)
        self.__log.debug(f"Exported mapping table to {file_name}.")

    def get_mapping_table(self, type: int) -> object:
        return self.com_obj.GetMappingTable(type)

    def get_mapping_table_by_name(self, type_name: str) -> object:
        return self.com_obj.GetMappingTableByName(type_name)

    def import_mapping_table(self, file_name: str) -> None:
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
        return self.com_obj.Count

    @property
    def paths(self) -> list:
        list_of_paths = []
        for index in range(1, self.count + 1):
            list_of_paths.append(self.com_obj.Item(index))
        return list_of_paths

    def add(self, source_file: str) -> object:
        return self.com_obj.Add(source_file)

    def clear(self) -> None:
        self.com_obj.Clear()

    def remove(self, index: int) -> None:
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
        return self.com_obj.End

    @end.setter
    def end(self, time: str) -> None:
        self.com_obj.End = time
        self.__log.debug(f"Time section end time set to {time}.")

    @property
    def start(self) -> str:
        return self.com_obj.Start

    @start.setter
    def start(self, time: str) -> None:
        self.com_obj.Start = time
        self.__log.debug(f"Time section start time set to {time}.")

    @property
    def type(self) -> int:
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
    """The SimulationSetup object represents the Simulation Setup of CANoe."""
    def __init__(self, conf_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(conf_com_obj.SimulationSetup)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe simulation setup: {str(e)}')

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
    """The ReplayCollection object represents the Replay Blocks of the CANoe application."""
    def __init__(self, sim_setup_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(sim_setup_com_obj.ReplayCollection)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe replay collection: {str(e)}')

    @property
    def count(self) -> int:
        return self.com_obj.Count

    def add(self, name: str) -> object:
        return self.com_obj.Add(name)

    def remove(self, index: int) -> None:
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
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(replay_block_com_obj)
        except Exception as e:
            self.__log.error(f'😡 Error initializing CANoe replay block: {str(e)}')

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def path(self) -> str:
        return self.com_obj.Path

    @path.setter
    def path(self, path: str):
        self.com_obj.Path = path

    def start(self):
        self.com_obj.Start()

    def stop(self):
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
        return self.com_obj.Count
