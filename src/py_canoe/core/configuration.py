from typing import TYPE_CHECKING, Iterable, Union
if TYPE_CHECKING:
    from py_canoe.core.application import Application
    from py_canoe.core.child_elements.measurement_setup import Logging, ExporterSymbol, Message
import os
import win32com.client

from py_canoe.core.child_elements.measurement_setup import MeasurementSetup
from py_canoe.core.child_elements.databases import Databases
from py_canoe.core.child_elements.replay_collection import ReplayCollection
from py_canoe.helpers.common import DoEventsUntil, logger, wait

TEST_MODULE_START_EVENT_TIMEOUT = 5  # seconds


class ConfigurationEvents:
    def __init__(self):
        self.CONFIGURATION_CLOSED = False
        self.SYSTEM_VARIABLES_DEFINITION_CHANGED = False

    def OnClose(self):
        self.CONFIGURATION_CLOSED = True

    def OnSystemVariablesDefinitionChanged(self):
        self.SYSTEM_VARIABLES_DEFINITION_CHANGED = True


class Configuration:
    """
    The Configuration object represents the active configuration.
    """
    def __init__(self, app: 'Application'):
        self.app = app
        self.bus_types = self.app.bus_types
        self.com_object = win32com.client.Dispatch(self.app.com_object.Configuration)
        self.configuration_events: ConfigurationEvents = win32com.client.WithEvents(self.com_object, ConfigurationEvents)
        self.configuration_test_setup = lambda: self.test_setup
        self.__test_setup_environments = self.configuration_test_setup().test_environments.fetch_all_test_environments()
        self.__test_modules = list()

    def fetch_test_modules(self):
        for te_name, te_inst in self.__test_setup_environments.items():
            for tm_name, tm_inst in te_inst.get_all_test_modules().items():
                self.__test_modules.append({'name': tm_name, 'object': tm_inst, 'environment': te_name})

    @property
    def comment(self) -> str:
        return self.com_object.Comment

    @property
    def full_name(self) -> str:
        return self.com_object.FullName

    @property
    def mode(self) -> int:
        return self.com_object.Mode

    @mode.setter
    def mode(self, type: int):
        self.com_object.Mode = type

    @property
    def modified(self) -> bool:
        return self.com_object.Modified

    @modified.setter
    def modified(self, modified: bool):
        self.com_object.Modified = modified

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def online_setup(self) -> 'MeasurementSetup':
        return MeasurementSetup(self.com_object.OnlineSetup)

    @property
    def path(self) -> str:
        return self.com_object.Path

    @property
    def read_only(self) -> bool:
        return self.com_object.ReadOnly

    @property
    def saved(self) -> bool:
        return self.com_object.Saved

    @property
    def test_setup(self) -> 'TestSetup':
        return TestSetup(self.com_object)

    def save(self) -> bool:
        try:
            if self.saved:
                logger.warning("⚠️ CANoe configuration is already saved.")
                return True
            self.com_object.Save()
            logger.info("📢 CANoe configuration saved 💾 successfully 🎉")
            return True
        except Exception as e:
            logger.error(f"❌ Error saving CANoe configuration: {e}")
            return False

    def save_as(self, path: str, major: int, minor: int, prompt_user: bool = False, create_dir: bool = True) -> bool:
        try:
            if create_dir:
                dir_path = os.path.dirname(path)
                if dir_path:
                    os.makedirs(dir_path, exist_ok=True)
                    logger.info(f'📂 Created directory {dir_path} for saving configuration')
            self.com_object.SaveAs(path, major, minor, prompt_user)
            logger.info(f"📢 CANoe configuration saved 💾 as {path} successfully 🎉")
            return True
        except Exception as e:
            logger.error(f"❌ Error saving CANoe configuration as '{path}': {e}")
            return False

    def get_can_bus_statistics(self, channel: int) -> dict:
        try:
            can_stat_obj = self.online_setup.bus_statistics.BusStatistic(self.bus_types['CAN'], channel)
            keys = [
                'BusLoad', 'ChipState', 'Error', 'ErrorTotal', 'Extended', 'ExtendedTotal',
                'ExtendedRemote', 'ExtendedRemoteTotal', 'Overload', 'OverloadTotal', 'PeakLoad',
                'RxErrorCount', 'Standard', 'StandardTotal', 'StandardRemote', 'StandardRemoteTotal',
                'TxErrorCount'
            ]
            can_bus_stat_info = {key.lower(): getattr(can_stat_obj, key) for key in keys}
            logger.info(f'📜 CAN bus channel ({channel}) statistics:')
            for key, value in can_bus_stat_info.items():
                logger.info(f"    {key}: {value}")
            return can_bus_stat_info
        except Exception as e:
            logger.error(f"❌ Error retrieving CAN bus statistics for channel {channel}: {e}")
            return {}

    def add_offline_source_log_file(self, absolute_log_file_path: str) -> bool:
        try:
            if not os.path.isfile(absolute_log_file_path):
                logger.error(f"❌ Error: Offline source log file '{absolute_log_file_path}' does not exist.")
                return False
            offline_sources_obj = self.com_object.OfflineSetup.Source.Sources
            offline_sources_files = [offline_sources_obj.Item(i) for i in range(1, offline_sources_obj.Count + 1)]
            file_already_added = any([file == absolute_log_file_path for file in offline_sources_files])
            if file_already_added:
                logger.warning(f"⚠️ Offline source log file '{absolute_log_file_path}' is already added.")
            else:
                offline_sources_obj.Add(absolute_log_file_path)
                logger.info(f'📢 File "{absolute_log_file_path}" added as offline source')
            return True
        except Exception as e:
            logger.error(f"❌ Error adding offline source log file '{absolute_log_file_path}': {e}")
            return False

    def set_replay_block_file(self, block_name: str, recording_file_path: str) -> bool:
        try:
            replay_collection_obj = ReplayCollection(self.com_object.SimulationSetup.ReplayCollection)
            replay_blocks_obj_dict = dict()
            for i in range(1, replay_collection_obj.count + 1):
                replay_block_obj = replay_collection_obj.item(i)
                replay_blocks_obj_dict[replay_block_obj.name] = replay_block_obj
            if block_name in replay_blocks_obj_dict:
                replay_blocks_obj_dict[block_name].path = recording_file_path
                logger.info(f"📢 Replay block path for '{block_name}' set to '{recording_file_path}'")
                return True
            else:
                logger.warning(f"⚠️ Replay block '{block_name}' not found")
                return False
        except Exception as e:
            logger.error(f"❌ Error setting replay block file for '{block_name}': {e}")
            return False

    def control_replay_block(self, block_name: str, start_stop: bool) -> bool:
        try:
            replay_collection_obj = ReplayCollection(self.com_object.SimulationSetup.ReplayCollection)
            replay_blocks_obj_dict = dict()
            for i in range(1, replay_collection_obj.count + 1):
                replay_block_obj = replay_collection_obj.item(i)
                replay_blocks_obj_dict[replay_block_obj.name] = replay_block_obj
            if block_name in replay_blocks_obj_dict:
                if start_stop:
                    replay_blocks_obj_dict[block_name].start()
                    logger.info(f"📢 Replay block '{block_name}' started")
                else:
                    replay_blocks_obj_dict[block_name].stop()
                    logger.info(f"📢 Replay block '{block_name}' stopped")
                return True
            else:
                logger.warning(f"⚠️ Replay block '{block_name}' not found")
                return False
        except Exception as e:
            logger.error(f"❌ Error controlling replay block '{block_name}': {e}")
            return False

    def enable_disable_replay_block(self, block_name: str, enable_disable: bool) -> bool:
        try:
            replay_collection_obj = ReplayCollection(self.com_object.SimulationSetup.ReplayCollection)
            replay_blocks_obj_dict = dict()
            for i in range(1, replay_collection_obj.count + 1):
                replay_block_obj = replay_collection_obj.item(i)
                replay_blocks_obj_dict[replay_block_obj.name] = replay_block_obj
            if block_name in replay_blocks_obj_dict:
                replay_blocks_obj_dict[block_name].enabled = enable_disable
                logger.info(f"📢 Replay block '{block_name}' {'enabled' if enable_disable else 'disabled'}")
                return True
            else:
                logger.warning(f"⚠️ Replay block '{block_name}' not found")
                return False
        except Exception as e:
            logger.error(f"❌ Error enabling/disabling replay block '{block_name}': {e}")
            return False

    def get_test_environments(self) -> dict:
        try:
            return self.__test_setup_environments
        except Exception as e:
            logger.error(f'❌ failed to get test environments. {e}')
            return {}

    def get_test_modules(self, env_name: str) -> dict:
        try:
            test_environments = self.get_test_environments()
            if len(test_environments) > 0:
                if env_name in test_environments.keys():
                    return test_environments[env_name].get_all_test_modules()
                else:
                    logger.warning(f'⚠️ "{env_name}" not found in configuration')
                    return {}
            else:
                logger.warning('⚠️ Zero test environments found in configuration. Not possible to fetch test modules')
                return {}
        except Exception as e:
            logger.error(f'❌ failed to get test modules. {e}')
            return {}

    def execute_test_module(self, test_module_name: str) -> int:
        try:
            test_verdict = {0: 'NotAvailable',
                            1: 'Passed',
                            2: 'Failed',
                            3: 'None (not available for test modules)',
                            4: 'Inconclusive (not available for test modules)',
                            5: 'ErrorInTestSystem (not available for test modules)', }
            execution_result = 0
            test_module_found = False
            test_env_name = ''
            for tm in self.__test_modules:
                if tm['name'] == test_module_name:
                    test_module_found = True
                    tm_obj = tm['object']
                    test_env_name = tm['environment']
                    logger.info(f'🔎 test module "{test_module_name}" found in "{test_env_name}"')
                    tm_obj.start()
                    tm_obj.wait_for_completion()
                    execution_result = tm_obj.verdict
                    break
                else:
                    continue
            if test_module_found and (execution_result == 1):
                logger.info(f'🧪🟢 test module "{test_env_name}.{test_module_name}" verdict = {test_verdict[execution_result]}')
            elif test_module_found and (execution_result != 1):
                logger.info(f'🧪🔴 test module "{test_env_name}.{test_module_name}" verdict = {test_verdict[execution_result]}')
            else:
                logger.warning(f'🧪⚠️ test module "{test_module_name}" not found. not possible to execute')
            return execution_result
        except Exception as e:
            logger.error(f'❌ failed to execute test module. {e}')
            return 0

    def stop_test_module(self, test_module_name: str):
        try:
            for tm in self.__test_modules:
                if tm['name'] == test_module_name:
                    tm['object'].stop()
                    test_env_name = tm['environment']
                    logger.info(f'🧪⏹️ test module "{test_module_name}" in test environment "{test_env_name}" stopped 🧍‍♂️')
            else:
                logger.warning(f'🧪⚠️ test module "{test_module_name}" not found. not possible to execute')
        except Exception as e:
            logger.error(f'❌ failed to stop test module. {e}')

    def execute_all_test_modules_in_test_env(self, env_name: str):
        try:
            test_modules = self.get_test_modules(env_name=env_name)
            if test_modules:
                for tm_name in test_modules.keys():
                    self.execute_test_module(tm_name)
            else:
                logger.warning(f'🧪⚠️ test modules not available in "{env_name}" test environment')
        except Exception as e:
            logger.error(f'🧪❌ failed to execute all test modules in "{env_name}" test environment. {e}')

    def stop_all_test_modules_in_test_env(self, env_name: str):
        try:
            test_modules = self.get_test_modules(env_name=env_name)
            if test_modules:
                for tm_name in test_modules.keys():
                    self.stop_test_module(tm_name)
            else:
                logger.warning(f'🧪⚠️ test modules not available in "{env_name}" test environment')
        except Exception as e:
            logger.error(f'🧪❌ failed to stop all test modules in "{env_name}" test environment. {e}')

    def execute_all_test_environments(self):
        try:
            test_environments = self.get_test_environments()
            if len(test_environments) > 0:
                for test_env_name in test_environments.keys():
                    logger.info(f'🧐🏃‍➡️🥱 executing test environment "{test_env_name}". please wait...')
                    self.execute_all_test_modules_in_test_env(test_env_name)
                    logger.info(f'🧐🧍✔️ completed executing test environment "{test_env_name}"')
            else:
                logger.warning('🧐⚠️ Zero test environments found in configuration')
        except Exception as e:
            logger.error(f'🧐❌ failed to execute all test environments. {e}')

    def stop_all_test_environments(self):
        try:
            test_environments = self.get_test_environments()
            if len(test_environments) > 0:
                for test_env_name in test_environments.keys():
                    logger.info(f'🧐⏹️🥱 stopping test environment "{test_env_name}" execution. please wait...')
                    self.stop_all_test_modules_in_test_env(test_env_name)
                    logger.info(f'🧐🧍✔️ completed stopping test environment "{test_env_name}"')
            else:
                logger.warning('🧐⚠️ Zero test environments found in configuration')
        except Exception as e:
            logger.error(f'🧐❌ failed to stop all test environments. {e}')

    def add_database(self, database_file: str, database_channel: int, database_network: Union[str, None]=None) -> bool:
        try:
            if self.app.measurement.running:
                logger.warning("⚠️ Cannot add database while measurement is running. Please stop the measurement first.")
                return False
            else:
                databases = Databases(self.com_object.GeneralSetup.DatabaseSetup.Databases)
                databases_info = {databases.item(index).full_name: databases.item(index) for index in range(1, databases.count + 1)}
                if database_file in databases_info.keys():
                    logger.warning(f'⚠️ database "{database_file}" already added')
                    return False
                else:
                    if database_network:
                        database = databases.add_network(database_file, database_network)
                    else:
                        database = databases.add(database_file)
                    wait(0.5)
                    database.channel = database_channel
                    wait(0.5)
                    logger.info(f'📢 database "{database_file}" added successfully to channel {database_channel}')
                    return True
        except Exception as e:
            logger.error(f"❌ Error adding database '{database_file}': {e}")
            return False

    def remove_database(self, database_file: str, database_channel: int) -> bool:
        try:
            if self.app.measurement.running:
                logger.warning("⚠️ Cannot remove database while measurement is running. Please stop the measurement first.")
                return False
            else:
                databases = Databases(self.com_object.GeneralSetup.DatabaseSetup.Databases)
                databases_info = {databases.item(index).full_name: databases.item(index) for index in range(1, databases.count + 1)}
                if database_file not in databases_info.keys():
                    logger.warning(f'⚠️ database "{database_file}" not available to remove')
                    return False
                else:
                    for index in range(1, databases.count + 1):
                        database = databases.item(index)
                        if (database.full_name == database_file) and (database.channel == database_channel):
                            databases.remove(index)
                            wait(1)
                            logger.info(f'📢 database "{database_file}" removed from channel "{database_channel}"')
                            return True
                    logger.warning(f'⚠️ unable to remove database "{database_file}" from channel "{database_channel}"')
                    return False
        except Exception as e:
            logger.error(f"❌ Error removing database '{database_file}': {e}")
            return False

    def get_mode(self) -> int:
        logger.info(f"⚙️ CANoe Configuration mode = ({self.mode} - {'Offline mode' if self.mode == 1 else 'Online mode'})")
        return self.mode

    def set_mode(self, type: int) -> bool:
        try:
            if type in [0, 1]:
                self.mode = type
                logger.info(f"⚙️ CANoe Configuration mode set to ({type} - {'Offline mode' if type == 1 else 'Online mode'})")
                return True
            else:
                logger.warning("⚠️ Invalid mode type. Use 0 for Offline mode and 1 for Online mode.")
                return False
        except Exception as e:
            logger.error(f"❌ Error setting CANoe Configuration mode: {e}")
            return False

    def get_logging_blocks(self) -> list['Logging']:
        blocks = []
        for i in range(1, self.online_setup.logging_collection.count + 1):
            logging_block = self.online_setup.logging_collection.item(i)
            blocks.append(logging_block)
        return blocks

    def add_logging_block(self, full_name: str) -> 'Logging':
        return self.online_setup.logging_collection.add(full_name)

    def remove_logging_block(self, index: int) -> None:
        if index == 0:
            raise ValueError("Logging blocks indexing starts from 1 and not 0.")
        self.online_setup.logging_collection.remove(index)

    def load_logs_for_exporter(self, logger_index: int) -> None:
        self.online_setup.logging_collection.item(logger_index).exporter.load()

    def get_symbols(self, logger_index: int) -> list['ExporterSymbol']:
        return self.online_setup.logging_collection.item(logger_index).exporter.symbols

    def get_messages(self, logger_index: int) -> list['Message']:
        return self.online_setup.logging_collection.item(logger_index).exporter.messages

    def add_filters_to_exporter(self, logger_index: int, full_names: 'Iterable'):
        expo_filter = self.online_setup.logging_collection.item(logger_index).exporter.filter
        for name in full_names:
            expo_filter.add(name)

    def start_export(self, logger_index: int):
        self.online_setup.logging_collection.item(logger_index).exporter.save()

    def start_stop_online_logging_block(self, full_name: str, start_stop: bool) -> bool:
        try:
            logging_blocks = self.get_logging_blocks()
            for logging_block in logging_blocks:
                if logging_block.full_name.lower() == full_name.lower():
                    if start_stop:
                        logging_block.trigger.start()
                        logger.info(f'📢 logging block {full_name} started')
                    else:
                        logging_block.trigger.stop()
                        logger.info(f'📢 logging block {full_name} stopped')
                    return True
            logger.warning(f'⚠️ logging block {full_name} not found.')
            return False
        except Exception as e:
            logger.error(f"❌ Error starting/stopping logging block {full_name}. {e}")

    def set_configuration_modified(self, modified: bool) -> None:
        self.modified = modified


class TestSetup:
    """The TestSetup object represents CANoe's test setup."""
    def __init__(self, conf_com_obj):
        self.com_object = win32com.client.Dispatch(conf_com_obj.TestSetup)

    def save_all(self, prompt_user=False) -> None:
        self.com_object.SaveAll(prompt_user)

    @property
    def test_environments(self) -> 'TestEnvironments':
        return TestEnvironments(self.com_object)


class TestEnvironments:
    """The TestEnvironments object represents the test environments within CANoe's test setup."""
    def __init__(self, test_setup_com_obj):
        self.com_object = win32com.client.Dispatch(test_setup_com_obj.TestEnvironments)

    @property
    def count(self) -> int:
        return self.com_object.Count

    def add(self, name: str) -> 'TestEnvironment':
        return TestEnvironment(self.com_object.Add(name))

    def remove(self, index: int, prompt_user=False) -> None:
        self.com_object.Remove(index, prompt_user)

    def fetch_all_test_environments(self) -> dict['str': 'TestEnvironment']:
        test_environments = dict()
        for index in range(1, self.count + 1):
            te_inst = TestEnvironment(self.com_object.Item(index))
            test_environments[te_inst.name] = te_inst
        return test_environments


class TestEnvironment:
    """The TestEnvironment object represents a test environment within CANoe's test setup."""
    def __init__(self, test_environment_com_obj):
        self.com_object = win32com.client.Dispatch(test_environment_com_obj)
        self.__test_modules = TestModules(self.com_object)
        self.__test_setup_folders = TestSetupFolders(self.com_object)
        self.__all_test_modules = {}
        self.__all_test_setup_folders = {}

    @property
    def enabled(self) -> bool:
        return self.com_object.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.com_object.Enabled = value

    @property
    def full_name(self) -> str:
        return self.com_object.FullName

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def path(self) -> str:
        return self.com_object.Path

    def execute_all(self) -> None:
        self.com_object.ExecuteAll()

    def save(self, name: str, prompt_user=False) -> None:
        self.com_object.Save(name, prompt_user)

    def save_as(self, name: str, major: int, minor: int, prompt_user=False) -> None:
        self.com_object.SaveAs(name, major, minor, prompt_user)

    def stop_sequence(self) -> None:
        self.com_object.StopSequence()

    def update_all_test_setup_folders(self, tsfs_instance=None):
        if tsfs_instance is None:
            tsfs_instance = self.__test_setup_folders
        if tsfs_instance.count > 0:
            test_setup_folders = tsfs_instance.fetch_test_setup_folders()
            for tsf_name, tsf_inst in test_setup_folders.items():
                self.__all_test_setup_folders[tsf_name] = tsf_inst
                if tsf_inst.test_modules.count > 0:
                    self.__all_test_modules.update(tsf_inst.test_modules.fetch_test_modules())
                if tsf_inst.folders.count > 0:
                    tsfs_instance = tsf_inst.folders
                    self.update_all_test_setup_folders(tsfs_instance)

    def get_all_test_modules(self):
        self.update_all_test_setup_folders()
        self.__all_test_modules.update(self.__test_modules.fetch_test_modules())
        return self.__all_test_modules


class TestSetupFolders:
    """The TestSetupFolders object represents the folders in a test environment or in a test setup folder."""
    def __init__(self, test_env_com_obj) -> None:
        self.com_object = test_env_com_obj.Folders

    @property
    def count(self) -> int:
        return self.com_object.Count

    def add(self, full_name: str) -> 'TestSetupFolderExt':
        return TestSetupFolderExt(self.com_object.Add(full_name))

    def remove(self, index: int, prompt_user=False) -> None:
        self.com_object.Remove(index, prompt_user)

    def fetch_test_setup_folders(self) -> dict:
        test_setup_folders = dict()
        for index in range(1, self.count + 1):
            tsf_com_obj = self.com_object.Item(index)
            tsf_inst = TestSetupFolderExt(tsf_com_obj)
            test_setup_folders[tsf_inst.name] = tsf_inst
        return test_setup_folders


class TestSetupFolderExt:
    """The TestSetupFolderExt object represents a directory in CANoe's test setup."""
    def __init__(self, test_setup_folder_ext_com_obj) -> None:
        self.com_object = win32com.client.Dispatch(test_setup_folder_ext_com_obj)

    @property
    def enabled(self) -> bool:
        return self.com_object.Enabled

    @enabled.setter
    def enabled(self, enabled: bool):
        self.com_object.Enabled = enabled

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def folders(self) -> 'TestSetupFolders':
        return TestSetupFolders(self.com_object)

    @property
    def test_modules(self) -> 'TestModules':
        return TestModules(self.com_object)

    def execute_all(self):
        self.com_object.ExecuteAll()

    def stop_sequence(self):
        self.com_object.StopSequence()


class TestModules:
    def __init__(self, test_env_com_obj) -> None:
        self.com_object = test_env_com_obj.TestModules

    @property
    def count(self) -> int:
        return self.com_object.Count

    def add(self, full_name: str) -> 'TestModule':
        return TestModule(self.com_object.Add(full_name))

    def remove(self, index: int, prompt_user=False) -> None:
        self.com_object.Remove(index, prompt_user)

    def fetch_test_modules(self) -> dict['str': 'TestModule']:
        test_modules = dict()
        for index in range(1, self.count + 1):
            tm_inst = TestModule(self.com_object.Item(index))
            test_modules[tm_inst.name] = tm_inst
        return test_modules


class TestModuleEvents:
    """test module events object."""
    def __init__(self):
        self.TM_STARTED = False
        self.TM_PAUSED = False
        self.TM_STOPPED = False
        self.TM_STOP_REASON = -1
        self.VALUE_TABLE_STOP_REASON = {
            0: "TestModuleEnd: The test module was executed completely",
            1: "UserAbortion: The test module was stopped by the user",
            2: "GeneralError: The test module was stopped by measurement stop"
        }
        self.TM_REPORT_GENERATED = False
        self.TEST_REPORT_INFORMATION = dict()
        self.TC_FAIL = False

    def OnStart(self):
        self.TM_STARTED = True

    def OnPause(self):
        self.TM_PAUSED = True

    def OnStop(self, reason):
        self.TM_STOP_REASON = reason
        self.TM_STOPPED = True

    def OnReportGenerated(self, success, sourceFullName, generatedFullName):
        self.TEST_REPORT_INFORMATION = {
            "success": success,
            "source_full_name": sourceFullName,
            "generated_full_name": generatedFullName
        }
        self.TM_REPORT_GENERATED = True

    def OnVerdictFail(self):
        self.TC_FAIL = True


class TestModule:
    """The TestModule object represents a test module in CANoe's test setup."""

    def __init__(self, test_module):
        self.com_object = win32com.client.Dispatch(test_module)
        self.test_module_events: TestModuleEvents = win32com.client.WithEvents(self.com_object, TestModuleEvents)
        self.VALUE_TABLE_VERDICT = {
            0: "NotAvailable",
            1: "Passed",
            2: "Failed",
            3: "None",
            4: "Inconclusive",
            5: "ErrorInTestSystem"
        }
        self.VALUE_TABLE_VERDICT_IMPACT = {
            0: "NoImpact",
            1: "EndTestCaseOnFail",
            2: "EndTestModuleOnFail"
        }

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def full_name(self) -> str:
        return self.com_object.FullName

    @property
    def path(self) -> str:
        return self.com_object.Path

    @property
    def number_of_executions(self) -> int:
        return self.com_object.NumberOfExecutions

    @number_of_executions.setter
    def number_of_executions(self, value: int):
        self.com_object.NumberOfExecutions = value

    @property
    def randomize_each_cycle(self) -> bool:
        return self.com_object.RandomizeEachCycle

    @randomize_each_cycle.setter
    def randomize_each_cycle(self, value: bool):
        self.com_object.RandomizeEachCycle = value

    @property
    def start_on_env_var(self) -> str:
        return self.com_object.StartOnEnvVar

    @start_on_env_var.setter
    def start_on_env_var(self, value: str):
        self.com_object.StartOnEnvVar = value

    @property
    def start_on_key(self) -> str:
        return self.com_object.StartOnKey

    @start_on_key.setter
    def start_on_key(self, value: str):
        self.com_object.StartOnKey = value

    @property
    def start_on_measurement(self) -> bool:
        return self.com_object.StartOnMeasurement

    @start_on_measurement.setter
    def start_on_measurement(self, value: bool):
        self.com_object.StartOnMeasurement = value

    @property
    def start_on_sys_var(self) -> str:
        return self.com_object.StartOnSysVar

    @start_on_sys_var.setter
    def start_on_sys_var(self, value: str):
        self.com_object.StartOnSysVar = value

    @property
    def test_cases_executed_in_random_order(self) -> bool:
        return self.com_object.TestCasesExecutedInRandomOrder

    @test_cases_executed_in_random_order.setter
    def test_cases_executed_in_random_order(self, value: bool):
        self.com_object.TestCasesExecutedInRandomOrder = value

    @property
    def test_state_sys_var(self) -> str:
        return self.com_object.TestStateSysVar

    @test_state_sys_var.setter
    def test_state_sys_var(self, value: str):
        self.com_object.TestStateSysVar = value

    @property
    def verdict(self) -> int:
        return self.com_object.Verdict

    @property
    def verdict_impact(self) -> int:
        return self.com_object.VerdictImpact

    @verdict_impact.setter
    def verdict_impact(self, value: int):
        self.com_object.VerdictImpact = value

    def _init_tm_event_variables(self):
        self.test_module_events.TM_STARTED = False
        self.test_module_events.TM_PAUSED = False
        self.test_module_events.TM_STOPPED = False
        self.test_module_events.TM_STOP_REASON = -1
        self.test_module_events.TM_REPORT_GENERATED = False
        self.test_module_events.TEST_REPORT_INFORMATION = dict()
        self.test_module_events.TC_FAIL = False

    def start(self):
        self._init_tm_event_variables()
        self.com_object.Start()
        status = DoEventsUntil(lambda: self.test_module_events.TM_STARTED, TEST_MODULE_START_EVENT_TIMEOUT, "Test Module Start")
        if status:
            logger.info(f'🧪🏃‍➡️ started executing test module ({self.name})...')

    def wait_for_completion(self) -> bool:
        return_value = False
        if self.test_module_events.TM_STARTED:
            logger.info(f'🧪🥱 waiting for test module ({self.name}) to complete...')
            while not self.test_module_events.TM_STOPPED:
                wait(0.01)
            logger.info(f'🧪🧍 test module ({self.name}) execution completed with stop reason 👉 {self.test_module_events.VALUE_TABLE_STOP_REASON[self.test_module_events.TM_STOP_REASON]}')
            return_value = True
        else:
            logger.warning(f'🧪⚠️ Test Module ({self.name}) is not started. Start the Test Module first.')
        return return_value

    def pause(self) -> bool:
        if self.test_module_events.TM_STARTED:
            self.com_object.Pause()
            logger.info(f'🧪🥱 pausing test module ({self.name}). please wait...')
            while not self.test_module_events.TM_PAUSED:
                wait(0.01)
            logger.info(f'🧪⏸️ paused test module ({self.name}).')
            return True
        else:
            logger.warning(f'🧪⚠️ Test Module ({self.name}) is not started. Start the Test Module first.')
            return False

    def resume(self) -> None:
        self.com_object.Resume()

    def stop(self) -> bool:
        if self.test_module_events.TM_STARTED:
            self.com_object.Stop()
            logger.info(f'🧪🥱 stopping test module ({self.name}). please wait...')
            while not self.test_module_events.TM_STOPPED:
                wait(0.01)
            logger.info(f'🧪⏹️ stopped test module ({self.name}).')
            return True
        else:
            logger.warning(f'🧪⚠️ Test Module ({self.name}) is not started. Start the Test Module first.')
            return False

    def reload(self) -> None:
        self.com_object.Reload()

    def set_execution_time(self, days: int, hours: int, minutes: int):
        self.com_object.SetExecutionTime(days, hours, minutes)
