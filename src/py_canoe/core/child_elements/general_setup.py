import win32com.client

from py_canoe.core.child_elements.ccp_setup import CCPSetup
from py_canoe.core.child_elements.channels import CanController
from py_canoe.core.child_elements.database_setup import DatabaseSetup
from py_canoe.core.child_elements.diagnostics_setup import DiagnosticsSetup
from py_canoe.core.child_elements.macros_setup import MacrosSetup
from py_canoe.core.child_elements.panel_setup import PanelSetup
from py_canoe.core.child_elements.security_setup import SecuritySetup
from py_canoe.core.child_elements.snippet_setup import SnippetSetup
from py_canoe.core.child_elements.visual_sequence_setup import VisualSequenceSetup
from py_canoe.core.child_elements.xcp_setup import XCPSetup


class GeneralSetup:
    """
    The MeasurementSetup object rRepresents the general settings of a CANoe configuration.
    """
    def __init__(self, general_setup_com_object) -> None:
        self.com_object = win32com.client.Dispatch(general_setup_com_object)

    @property
    def ccp_setup(self) -> 'CCPSetup':
        return CCPSetup(self.com_object.CCPSetup)

    def get_channels_count(self, bust_type: int) -> int:
        return self.com_object.Channels(bust_type)

    def set_channels_count(self, bust_type: int, channel: int):
        self.com_object.SetChannels(bust_type, channel)

    @property
    def controller_setup(self, bust_type: int, channel: int) -> 'CanController':
        return CanController(self.com_object.ControllerSetup(bust_type, channel))

    @property
    def database_setup(self) -> 'DatabaseSetup':
        return DatabaseSetup(self.com_object.DatabaseSetup)

    @property
    def diagnostics_setup(self) -> 'DiagnosticsSetup':
        return DiagnosticsSetup(self.com_object.DiagnosticsSetup)

    @property
    def macros_setup(self) -> 'MacrosSetup':
        return MacrosSetup(self.com_object.MacrosSetup)

    @property
    def panel_setup(self) -> 'PanelSetup':
        return PanelSetup(self.com_object.PanelSetup)

    @property
    def security_setup(self) -> 'SecuritySetup':
        return SecuritySetup(self.com_object.SecuritySetup)

    @property
    def snippet_setup(self) -> 'SnippetSetup':
        return SnippetSetup(self.com_object.SnippetSetup)

    @property
    def visual_sequence_setup(self) -> 'VisualSequenceSetup':
        return VisualSequenceSetup(self.com_object.VisualSequenceSetup)

    @property
    def xcp_setup(self) -> 'XCPSetup':
        return XCPSetup(self.com_object.XCPSetup)
