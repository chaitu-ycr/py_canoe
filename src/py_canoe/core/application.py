from typing import Union
import pythoncom
import win32com.client
import win32com.client.gencache

from py_canoe.utils.common import DoEventsUntil
from py_canoe.utils.common import logger


class ApplicationEvents:
    CONFIGURATION_OPENED: bool = False
    CANOE_IS_QUIT: bool = False

    @staticmethod
    def OnOpen(fullname):
        ApplicationEvents.CONFIGURATION_OPENED = True
        logger.info(f'[EVENT][APPLICATION] 📢 CANoe Configuration Opened: {fullname} 🎉')

    @staticmethod
    def OnQuit():
        ApplicationEvents.CANOE_IS_QUIT = True
        logger.info('[EVENT][APPLICATION] 📢 CANoe Application Quit 🎉')


def wait_for_event_canoe_configuration_to_open(timeout: Union[int, float]) -> bool:
    ApplicationEvents.CONFIGURATION_OPENED = False
    status = DoEventsUntil(lambda: ApplicationEvents.CONFIGURATION_OPENED, timeout, "CANoe Configuration Open")
    if not status:
        logger.error(f"😡 Unable to open CANoe configuration within {timeout} seconds.")
    return status

def wait_for_event_canoe_quit(timeout: Union[int, float]) -> bool:
    ApplicationEvents.CANOE_IS_QUIT = False
    status = DoEventsUntil(lambda: ApplicationEvents.CANOE_IS_QUIT, timeout, "CANoe Application Quit")
    if not status:
        logger.error(f"😡 Unable to quit CANoe application within {timeout} seconds.")
    return status

def new(app, auto_save: bool = False, prompt_user: bool = False, timeout=5) -> bool:
    try:
        if app.com_object is None:
            pythoncom.CoInitialize()
            app.com_object = win32com.client.gencache.EnsureDispatch("CANoe.Application")
        win32com.client.WithEvents(app.com_object, ApplicationEvents)
        app.com_object.New(auto_save, prompt_user)
        status = wait_for_event_canoe_configuration_to_open(timeout)
        if status:
            logger.info('📢 New empty CANoe configuration Opened 🎉')
        return status
    except Exception as e:
        logger.error(f"😡 Error creating new CANoe configuration: {e}")
        return False

def open(app, canoe_cfg: str, visible=True, auto_save=True, prompt_user=False, auto_stop=True, timeout=5) -> bool:
    try:
        if app.com_object is None:
            pythoncom.CoInitialize()
            app.com_object = win32com.client.gencache.EnsureDispatch("CANoe.Application")
            win32com.client.WithEvents(app.com_object, ApplicationEvents)
            win32com.client.WithEvents(app.com_object.Measurement, app.measurement.MeasurementEvents)
            win32com.client.WithEvents(app.com_object.Configuration, app.configuration.ConfigurationEvents)
        app.com_object.Visible = visible
        if auto_stop:
            app.stop_measurement(timeout=timeout)
        app.com_object.Open(canoe_cfg, auto_save, prompt_user)
        status = wait_for_event_canoe_configuration_to_open(timeout)
        if status:
            app._fetch_diagnostic_devices()
        return status
    except Exception as e:
        logger.error(f"😡 Error opening CANoe configuration '{canoe_cfg}': {e}")
        return False

def quit(app, timeout=5) -> bool:
    try:
        if app.com_object is None:
            logger.warning("⚠️ Cannot quit, CANoe COM object is not initialized.")
            return False
        else:
            app.com_object.Configuration.Modified = False
            app.com_object.Quit()
            status = wait_for_event_canoe_quit(timeout)
            if status:
                # pythoncom.CoUninitialize()
                # app.com_object = None
                return True
            else:
                return False
    except Exception as e:
        logger.error(f"😡 Error quitting CANoe application: {e}")
        return False

def get_running_instance(app, visible=True) -> Union[win32com.client.CDispatch, None]:
    try:
        if app.com_object is None:
            pythoncom.CoInitialize()
            app.com_object = win32com.client.gencache.EnsureDispatch("CANoe.Application")
            win32com.client.WithEvents(app.com_object, ApplicationEvents)
            win32com.client.WithEvents(app.com_object.Measurement, app.measurement.MeasurementEvents)
            win32com.client.WithEvents(app.com_object.Configuration, app.configuration.ConfigurationEvents)
        app.com_object.Visible = visible
        app._fetch_diagnostic_devices()
        return app.com_object
    except Exception as e:
        logger.error(f"😡 Error fetching running instance of CANoe application: {e}")
        return None
