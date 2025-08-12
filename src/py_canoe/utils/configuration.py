import os
import win32com.client
from typing import Union
from py_canoe.utils.common import logger
from py_canoe.utils.common import DoEventsUntil


class ConfigurationEvents:
    CONFIGURATION_CLOSED: bool = False
    SYSTEM_VARIABLES_DEFINITION_CHANGED: bool = False

    @staticmethod
    def OnClose():
        logger.info('[EVENT][CONFIGURATION] 📢 CANoe Configuration Closed')
        ConfigurationEvents.CONFIGURATION_CLOSED = True

    @staticmethod
    def OnSystemVariablesDefinitionChanged():
        logger.info('[EVENT][CONFIGURATION] 📢 CANoe System Variables Definition Changed')
        ConfigurationEvents.SYSTEM_VARIABLES_DEFINITION_CHANGED = True


def wait_for_event_canoe_configuration_closed(timeout: Union[int, float]) -> bool:
    ConfigurationEvents.CONFIGURATION_CLOSED = False
    status = DoEventsUntil(lambda: ConfigurationEvents.CONFIGURATION_CLOSED, timeout, "CANoe Configuration Close")
    if not status:
        logger.error(f"😡 Error: CANoe configuration did not close within {timeout} seconds.")
    return status

def wait_for_event_canoe_system_variables_definition_changed(timeout: Union[int, float]) -> bool:
    ConfigurationEvents.SYSTEM_VARIABLES_DEFINITION_CHANGED = False
    status = DoEventsUntil(lambda: ConfigurationEvents.SYSTEM_VARIABLES_DEFINITION_CHANGED, timeout, "CANoe System Variables Definition Change")
    if not status:
        logger.error(f"😡 Error: CANoe system variables definition did not change within {timeout} seconds.")
    return status

def save_configuration(app) -> bool:
    try:
        if app.com_object.Configuration.Saved:
            logger.warning("⚠️ CANoe configuration is already saved.")
            return True
        app.com_object.Configuration.Save()
        logger.info("📢 CANoe configuration saved 💾 successfully 🎉")
        return True
    except Exception as e:
        logger.error(f"😡 Error saving CANoe configuration: {e}")
        return False

def save_configuration_as(app, path: str, major: int, minor: int, prompt_user: bool = False, create_dir: bool = True) -> bool:
    try:
        if create_dir:
            dir_path = os.path.dirname(path)
            if dir_path:
                os.makedirs(dir_path, exist_ok=True)
                logger.info(f'📂 Created directory {dir_path} for saving configuration')
        app.com_object.Configuration.SaveAs(path, major, minor, prompt_user)
        logger.info(f"📢 CANoe configuration saved 💾 as {path} successfully 🎉")
        return True
    except Exception as e:
        logger.error(f"😡 Error saving CANoe configuration as '{path}': {e}")
        return False

def get_can_bus_statistics(app, channel: int) -> dict:
    try:
        can_stat_obj = app.com_object.Configuration.OnlineSetup.BusStatistics.BusStatistic(app.bus_type['CAN'], channel)
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
        logger.error(f"😡 Error retrieving CAN bus statistics for channel {channel}: {e}")
        return {}

def add_offline_source_log_file(app, absolute_log_file_path: str) -> bool:
    try:
        if not os.path.isfile(absolute_log_file_path):
            logger.error(f"😡 Error: Offline source log file '{absolute_log_file_path}' does not exist.")
            return False
        offline_sources_obj = app.com_object.Configuration.OfflineSetup.Source.Sources
        offline_sources_files = [offline_sources_obj.Item(i) for i in range(1, offline_sources_obj.Count + 1)]
        file_already_added = any([file == absolute_log_file_path for file in offline_sources_files])
        if file_already_added:
            logger.warning(f"⚠️ Offline source log file '{absolute_log_file_path}' is already added.")
        else:
            offline_sources_obj.Add(absolute_log_file_path)
            logger.info(f'📢 File "{absolute_log_file_path}" added as offline source')
        return True
    except Exception as e:
        logger.error(f"😡 Error adding offline source log file '{absolute_log_file_path}': {e}")
        return False

def set_replay_block_file(app, block_name: str, recording_file_path: str) -> bool:
    try:
        replay_collection_obj = app.com_object.Configuration.SimulationSetup.ReplayCollection
        replay_blocks_obj_dict = dict()
        for i in range(1, replay_collection_obj.Count + 1):
            replay_block_obj = win32com.client.Dispatch(replay_collection_obj.Item(i))
            replay_blocks_obj_dict[replay_block_obj.Name] = replay_block_obj
        if block_name in replay_blocks_obj_dict:
            replay_blocks_obj_dict[block_name].Path = recording_file_path
            logger.info(f"📢 Replay block path for '{block_name}' set to '{recording_file_path}'")
            return True
        else:
            logger.warning(f"⚠️ Replay block '{block_name}' not found")
            return False
    except Exception as e:
        logger.error(f"😡 Error setting replay block file for '{block_name}': {e}")
        return False

def control_replay_block(app, block_name: str, start_stop: bool) -> bool:
    try:
        replay_collection_obj = app.com_object.Configuration.SimulationSetup.ReplayCollection
        replay_blocks_obj_dict = dict()
        for i in range(1, replay_collection_obj.Count + 1):
            replay_block_obj = win32com.client.Dispatch(replay_collection_obj.Item(i))
            replay_blocks_obj_dict[replay_block_obj.Name] = replay_block_obj
        if block_name in replay_blocks_obj_dict:
            if start_stop:
                replay_blocks_obj_dict[block_name].Start()
                logger.info(f"📢 Replay block '{block_name}' started")
            else:
                replay_blocks_obj_dict[block_name].Stop()
                logger.info(f"📢 Replay block '{block_name}' stopped")
            return True
        else:
            logger.warning(f"⚠️ Replay block '{block_name}' not found")
            return False
    except Exception as e:
        logger.error(f"😡 Error controlling replay block '{block_name}': {e}")
        return False

def enable_disable_replay_block(app, block_name: str, enable_disable: bool) -> bool:
    try:
        replay_collection_obj = app.com_object.Configuration.SimulationSetup.ReplayCollection
        replay_blocks_obj_dict = dict()
        for i in range(1, replay_collection_obj.Count + 1):
            replay_block_obj = win32com.client.Dispatch(replay_collection_obj.Item(i))
            replay_blocks_obj_dict[replay_block_obj.Name] = replay_block_obj
        if block_name in replay_blocks_obj_dict:
            replay_blocks_obj_dict[block_name].Enabled = enable_disable
            logger.info(f"📢 Replay block '{block_name}' {'enabled' if enable_disable else 'disabled'}")
            return True
        else:
            logger.warning(f"⚠️ Replay block '{block_name}' not found")
            return False
    except Exception as e:
        logger.error(f"😡 Error enabling/disabling replay block '{block_name}': {e}")
        return False
