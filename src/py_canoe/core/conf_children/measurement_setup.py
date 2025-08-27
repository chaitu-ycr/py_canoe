from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from py_canoe.core.configuration import Configuration

import os
import win32com.client

from py_canoe.utils.common import DoEventsUntil, logger, wait


class Trigger:
    """
    The Trigger object represents the trigger block that is located before the Logging Block in the Measurement Setup.
    """
    def __init__(self, trigger_com):
        self.com_object = win32com.client.Dispatch(trigger_com)

    @property
    def active(self) -> bool:
        return self.com_object.Active

    @active.setter
    def active(self, value: bool):
        self.com_object.Active = value

    def start(self):
        self.com_object.Start()

    def stop(self):
        self.com_object.Stop()


class ExporterSymbol:
    """
    The ExporterSymbol object represents a symbol (signal, system variable or bus statistics information), found in a source file, loaded by the Exporter.
    """
    def __init__(self, message_com):
        self.com_object = win32com.client.Dispatch(message_com)

    @property
    def full_name(self) -> str:
        return self.com_object.FullName


class ExporterSymbols:
    """
    The ExporterSymbols object represents a collection of signals, system variables and bus statistics information, found in source files, loaded by the Exporter.
    """
    def __init__(self, symbols_com):
        self.com_object = win32com.client.Dispatch(symbols_com)

    @property
    def count(self) -> int:
        return int(self.com_object.Count)

    def item(self, index: int) -> 'ExporterSymbol':
        return ExporterSymbol(self.com_object.Item(index))


class Message:
    """
    The Message object represents a single message
    """
    def __init__(self, message_com):
        self.com_object = win32com.client.Dispatch(message_com)

    @property
    def full_name(self) -> str:
        return self.com_object.FullName


class Messages:
    """
    The Messages object represents a collection of messages.
    """
    def __init__(self, messages_com):
        self.com_object = win32com.client.Dispatch(messages_com)

    @property
    def count(self) -> int:
        return int(self.com_object.Count)

    def item(self, index: int) -> 'Message':
        return Message(self.com_object.Item(index))


class Filter:
    """
    The Filter object represents a Pass Filter for messages and signals in usage with an exporter.
    """
    def __init__(self, filter_com):
        self.com_object = win32com.client.Dispatch(filter_com)

    @property
    def count(self) -> int:
        return int(self.com_object.Count)

    @property
    def enabled(self) -> bool:
        return self.com_object.Enabled

    @enabled.setter
    def enabled(self, enabled: bool):
        self.com_object.Enabled = enabled

    def item(self, index):
        raise NotImplementedError("Item access is not implemented yet.")

    def add(self, fullname: str):
        self.com_object.Add(fullname)

    def clear(self):
        self.com_object.Clear()

    def remove(self, index: int):
        self.com_object.Remove(index)


class Exporter:
    """
    The Exporter object represents an export dialog, as it can be used in CANoe e.g. in a Logging Block in the Measurement Setup.
    """
    def __init__(self, exporter_com):
        self.com_object = win32com.client.Dispatch(exporter_com)

    def destinations(self):
        raise NotImplementedError("Destinations access is not implemented yet.")

    @property
    def filter(self) -> 'Filter':
        return Filter(self.com_object.Filter)

    @property
    def messages(self) -> list['Message']:
        messages_collection = Messages(self.com_object.Symbols)
        messages = []
        for i in range(1, messages_collection.count + 1):
            messages.append(messages_collection.item(i))
        return messages

    def settings(self):
        raise NotImplementedError("FileNameOptions access is not implemented yet.")

    def sources(self):
        raise NotImplementedError("FileNameOptions access is not implemented yet.")

    @property
    def symbols(self) -> list['ExporterSymbol']:
        symbols_collection = ExporterSymbols(self.com_object.Symbols)
        symbols = []
        for i in range(1, symbols_collection.count + 1):
            symbols.append(symbols_collection.item(i))
        return symbols

    def time_section(self):
        raise NotImplementedError("TimeSection access is not implemented yet.")

    def load(self):
        self.com_object.Load()

    def save(self, no_prompt_user: bool = True):
        self.com_object.Save(noPromptUser=no_prompt_user)


class Logging:
    """
    The Logging object represents a Logging Block in the Measurement Setup.
    """
    def __init__(self, logging_com):
        self.com_object = win32com.client.Dispatch(logging_com)

    @property
    def exporter(self) -> 'Exporter':
        return Exporter(self.com_object.Exporter)

    def file_name_options(self):
        raise NotImplementedError("FileNameOptions access is not implemented yet.")

    def filter(self):
        raise NotImplementedError("FileNameOptions access is not implemented yet.")

    @property
    def full_name(self) -> str:
        return self.com_object.FullName

    @full_name.setter
    def full_name(self, fullname: str):
        self.com_object.FullName = fullname

    @property
    def trigger(self) -> "Trigger":
        return Trigger(self.com_object.Trigger)


class LoggingCollection:
    """
    The LoggingCollection object is a collection of all Logging Blocks belonging to a Measurement Setup
    """
    def __init__(self, logging_collection_com):
        self.com_object = win32com.client.Dispatch(logging_collection_com)

    @property
    def count(self) -> int:
        return int(self.com_object.Count)

    def item(self, index: int) -> 'Logging':
        return Logging(self.com_object.Item(index))

    def add(self, full_name: str) -> 'Logging':
        return Logging(self.com_object.Add(full_name))

    def remove(self, index: int):
        self.com_object.Remove(index)


class MeasurementSetup:
    """
    The MeasurementSetup object represents the Measurement Setup in CANoe.
    """
    def __init__(self, meas_com_object) -> None:
        self.com_object = win32com.client.Dispatch(meas_com_object)

    @property
    def animation_factor(self):
        return self.com_object.AnimationFactor

    @animation_factor.setter
    def animation_factor(self, value: int):
        self.com_object.AnimationFactor = value

    @property
    def bus_statistics(self):
        return self.com_object.BusStatistics

    @property
    def logging_collection(self):
        return LoggingCollection(self.com_object.LoggingCollection)

    @property
    def offline_source_root(self):
        return self.com_object.OfflineSourceRoot

    @property
    def parallelization_level(self) -> int:
        return self.com_object.ParallelizationLevel

    @parallelization_level.setter
    def parallelization_level(self, level: int):
        self.com_object.ParallelizationLevel = level

    @property
    def source(self):
        return self.com_object.Source

    @property
    def video_windows(self):
        return self.com_object.VideoWindows

    @property
    def view_synchronization(self):
        return self.com_object.ViewSynchronization

    @property
    def working_mode(self) -> int:
        return self.com_object.WorkingMode

    @working_mode.setter
    def working_mode(self, mode: int):
        self.com_object.WorkingMode = mode
