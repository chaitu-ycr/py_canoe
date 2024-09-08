# import external modules here
import logging
from re import S
import re
import win32com.client
# import internal modules here

class CanoeConfigurationEvents:
    """Handler for CANoe Configuration events"""

    @staticmethod
    def OnClose():
        """Occurs when the configuration is closed.
        """
        logging.getLogger('CANOE_LOG').info('👉 configuration OnClose event triggered.')

    @staticmethod
    def OnSystemVariablesDefinitionChanged():
        """Occurs when system variable definitions are added, changed or removed.
        """
        logging.getLogger('CANOE_LOG').info('👉 configuration OnSystemVariablesDefinitionChanged event triggered.')

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
        self.__log.info(f"Configuration modified property value set to {value}.")

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
            bool: False is returned, If changes were made to the configuration and not been saved yet. otherwise True is returned.
        """
        return self.com_obj.Saved

    @property
    def offline_setup(self):
        return OfflineSetup(self.com_obj)

    @property
    def online_setup(self):
        return OnlineSetup(self.com_obj)

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
            self.log.info(f'Saved configuration({path}).')
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
        self.log.info(f'Saved configuration as {path}.')
        return self.saved


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
        self.__log.info(f"Exported mapping table to {file_name}.")

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
        self.__log.info(f"Imported mapping table from {file_name}.")

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
        self.__log.info(f"Time section end time set to {time}.")

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
        self.__log.info(f"Time section start time set to {time}.")

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

