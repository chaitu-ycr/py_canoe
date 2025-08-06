from typing import Union
from py_canoe.utils.common import logger
from py_canoe.utils.common import DoEventsUntil


class ApplicationEvents:
    CONFIGURATION_OPENED: bool = False
    CANOE_IS_QUIT: bool = False

    @staticmethod
    def OnOpen(fullname):
        ApplicationEvents.CONFIGURATION_OPENED = True
        logger.info(f'[EVENT][APPLICATION] CANoe Configuration Opened: {fullname}')

    @staticmethod
    def OnQuit():
        ApplicationEvents.CANOE_IS_QUIT = True
        logger.info('[EVENT][APPLICATION] CANoe Application Quit')


def wait_for_event_canoe_configuration_to_open(timeout: Union[int, float]) -> bool:
    ApplicationEvents.CONFIGURATION_OPENED = False
    status = DoEventsUntil(lambda: ApplicationEvents.CONFIGURATION_OPENED, timeout, "CANoe Configuration Open")
    if not status:
        logger.error(f"😡 Error: CANoe configuration did not open within {timeout} seconds.")
    return status

def wait_for_event_canoe_quit(timeout: Union[int, float]) -> bool:
    ApplicationEvents.CANOE_IS_QUIT = False
    status = DoEventsUntil(lambda: ApplicationEvents.CANOE_IS_QUIT, timeout, "CANoe Application Quit")
    if not status:
        logger.error(f"😡 Error: CANoe application did not quit within {timeout} seconds.")
    return status
