"""CANoe COM objects related to logging setup"""

import logging

import win32com

logger = logging.getLogger('CANOE_LOG')


class LoggingCollection:
    """Collection of all Logging Blocks of the current configuration."""

    def __init__(self, logging_collection_com):
        self._com = win32com.client.Dispatch(logging_collection_com)

    @property
    def count(self) -> int:
        """This property returns the number of objects inside the collection."""
        return int(self._com.Count)

    def item(self, index: int) -> "Logging":
        """This property returns an object from the collection."""
        return Logging(self._com.Item(index))

    def add(self, full_name: str) -> "Logging":
        """This method adds a logging block to the Measurement Setup.

        :param full_name: full path to log file as "C:/file.(asc|blf|mf4)", may have
                          field functions like {IncMeasurement} in the file name
        """
        return Logging(self._com.Add(full_name))

    def remove(self, index: int):
        """This method removes a logging block (Logging) from a logging collection."""
        self._com.Remove(index)


class Logging:

    """The Logging object represents a Logging Block in the current configuration.

    The filename extension of the logging file which is specified with this property
    determines the logging format. The full name can now also contain field functions;
    e.g. Logging.FullName = "LOGFILE_M{IncMeasurement}.ASC

    """

    def __init__(self, logging_com):
        self._com = win32com.client.Dispatch(logging_com)

    def exporter(self) -> "Exporter":
        """This property returns an Exporter object."""
        return Exporter(self._com.Exporter)

    def file_name_options(self):
        """This property returns a LoggingFileNameOptions object."""
        raise NotImplementedError("FileNameOptions access is not implemented yet.")

    def filter(self):
        """This property returns a LoggingFilter object."""
        raise NotImplementedError("FileNameOptions access is not implemented yet.")

    @property
    def full_name(self) -> str:
        """This property sets or determines the complete path to the logging file."""
        return self._com.FullName

    @full_name.setter
    def full_name(self, fullname: str):
        self._com.FullName = fullname

    def trigger(self) -> "Trigger":
        """This property returns a Trigger object."""
        return Trigger(self._com.Trigger)


class Trigger:

    """Trigger block located before the Logging Block in the Measurement Setup."""

    def __init__(self, trigger_com):
        self._com = win32com.client.Dispatch(trigger_com)

    @property
    def active(self) -> bool:
        """This property sets or returns the status of the trigger."""
        return self._com.Active

    @active.setter
    def active(self, _active: bool):
        self._com.Active = _active

    def start(self):
        """This method starts the trigger."""
        self._com.Start()

    def stop(self):
        """This method stops the trigger."""
        self._com.Stop()


class Exporter:

    """Export dialog.

    The Exporter object represents an export dialog, as it can
    be used e.g. in a Logging Block in the Measurement Setup.

    """

    def __init__(self, exporter_com):
        self._com = win32com.client.Dispatch(exporter_com)

    def destinations(self):
        """Returns a Files object for the destination files of an exporter.

        On the first access to the destination files of an exporter the Files object
        contains the destination file as it is defined in the according Export dialog.

        If no destination file is set a destination file is derived from the source
        file. Name and path of this destination file correspond to the name and path
        of the source file. The format is CSV. If as many destination files as source
        files are given, an according destination file will be generated for each
        source file (n to n).

        If the numbers of destination files and source files don't match, the number of
        destination files must be 1. All source files will be merged into the
        destination file (n to 1).
        """
        raise NotImplementedError("Destinations access is not implemented yet.")

    def filter(self) -> "Filter":
        """Returns a Filter object"""
        return Filter(self._com.Filter)

    def messages(self) -> list["Message"]:
        """Returns all messages that have been detected during loading.

        Consist of messages that come from the source files of an exporter
        and can be exported/converted.

        """
        messages_collection = Messages(self._com.Symbols)
        messages = []
        for i in range(1, messages_collection.count + 1):
            messages.append(messages_collection.item(i))
        return messages

    def settings(self):
        """Returns an ExporterSettings object."""
        raise NotImplementedError("FileNameOptions access is not implemented yet.")

    def sources(self):
        """Returns Files object for each source file of the exporter or offline source.

        On the first access to the source files of the exporter the Files object
        contains the source file as defined in the Export dialog.

        Typically this is the logging file as set in the Logging Block, as long as no
        other source file has been defined manually.

        """
        raise NotImplementedError("FileNameOptions access is not implemented yet.")

    def symbols(self) -> list["ExporterSymbol"]:
        """Returns all symbols that have been detected during loading the source files.

        This includes signals, system variables and bus statistics information
        that can be exported/converted.

        """
        symbols_collection = ExporterSymbols(self._com.Symbols)
        symbols = []
        for i in range(1, symbols_collection.count + 1):
            symbols.append(symbols_collection.item(i))
        return symbols

    def time_section(self):
        """Returns the TimeSection object."""
        raise NotImplementedError("TimeSection access is not implemented yet.")

    def load(self):
        """Loads source files of an exporter and determines the signals and messages.

        If several source files are set, all signals and messages of all
        source files are determined.

        """
        self._com.Load()

    def save(self, no_prompt_user: bool = True):
        """Starts the export/conversion.

        Although the parameter noPromptUser is set to True, the function will fail
        if a failure situation occurs and the storage cannot be performed. Possible
        failure situations are e.g. write-protection or disk full.

        """
        self._com.Save(noPromptUser=no_prompt_user)


class Filter:
    """Represents a Pass Filter for messages and signals in usage with an exporter."""

    def __init__(self, filter_com):
        self._com = win32com.client.Dispatch(filter_com)

    @property
    def count(self) -> int:
        """This property returns the number of objects inside the collection."""
        return int(self._com.Count)

    @property
    def enabled(self) -> bool:
        """This property activates/deactivates a Filter object or returns its state.

        The initial value of this property is False.

        """
        return self._com.Enabled

    @enabled.setter
    def enabled(self, enabled: bool):
        self._com.Enabled = enabled

    def item(self, index):
        """This property returns an object from the collection."""
        raise NotImplementedError("Item access is not implemented yet.")

    def add(self, fullname: str):
        """This method adds a message or a signal to the explorer's filter.

        The Exporter object needs fully qualified names of all messages and signals
        that have to be taken into consideration during export or conversion:

            Messages:
            <DatabaseName>::<MessageName>
            Signals:
            <DatabaseName>::<MessageName>::<SignalName>
            System variables:
            <Namespace>::<SystemVariable>
            Environment variables:
            <DatabaseName>::<EnvironmentVariable>

        When performing a message-oriented conversion only messages, system variables
        and environment variables are taken into consideration. Any (even specified)
        signals will be ignored.

        When performing a signal-orientierted export messages, signals, system variables
        and environment variables are taken into consideration. All signals included in
        specified messages, system variables and environment variables will be
        exported as well.

        Besides using their symbolic name messages can be declared by using their
        numeric ID, which either can be decimal or hexadecimal.

        :param fullname: of the signal or message

        """
        self._com.Add(fullname)

    def clear(self):
        """This method clears all messages and signals of the filter."""
        self._com.Clear()

    def remove(self, index: int):
        """This method removes a message or a signal from the filter.

        :param index: number

        """
        self._com.Remove(index)


class Messages:

    """Collection of messages."""

    def __init__(self, messages_com):
        self._com = win32com.client.Dispatch(messages_com)

    @property
    def count(self) -> int:
        """This property returns the number of objects inside the collection."""
        return int(self._com.Count)

    def item(self, index: int) -> "Message":
        """This property returns an object from the collection."""
        return Message(self._com.Item(index))


class ExporterSymbols:

    """Collection of signals."""

    def __init__(self, symbols_com):
        self._com = win32com.client.Dispatch(symbols_com)

    @property
    def count(self) -> int:
        """This property returns the number of objects inside the collection."""
        return int(self._com.Count)

    def item(self, index: int) -> "ExporterSymbol":
        """This property returns an object from the collection."""
        return ExporterSymbol(self._com.Item(index))


class Message:

    """The Message object represents a single message."""

    def __init__(self, message_com):
        self._com = win32com.client.Dispatch(message_com)

    def full_name(self) -> str:
        """This property determines the fully qualified name of a message.

        The following format is used: <DatabaseName>::<MessageName>

        """
        return self._com.FullName


class ExporterSymbol:

    """Symbol (signal, system variable or bus statistics information).

     Found in a source file, loaded by the Exporter.

     """

    def __init__(self, message_com):
        self._com = win32com.client.Dispatch(message_com)

    def full_name(self) -> str:
        """This property returns the fully qualified symbol name."""
        return self._com.FullName
