from typing import Union
from py_canoe.utils.common import logger
from py_canoe.utils.common import DoEventsUntil


class ConfigurationEvents:
    CONFIGURATION_CLOSED: bool = False
    SYSTEM_VARIABLES_DEFINITION_CHANGED: bool = False

    @staticmethod
    def OnClose():
        logger.info('[EVENT][CONFIGURATION] CANoe Configuration Closed')
        ConfigurationEvents.CONFIGURATION_CLOSED = True

    @staticmethod
    def OnSystemVariablesDefinitionChanged():
        logger.info('[EVENT][CONFIGURATION] CANoe System Variables Definition Changed')
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
