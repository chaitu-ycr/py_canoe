# import external modules here
import logging
import win32com.client


class Version:
    """The Version object represents the version of the CANoe application."""
    def __init__(self, app_com_obj):
        try:
            self.__log = logging.getLogger('CANOE_LOG')
            self.com_obj = win32com.client.Dispatch(app_com_obj.Version)
        except Exception as e:
            self.__log.error(f"Error while creating Version object: {e}")

    @property
    def build(self) -> int:
        return self.com_obj.Build

    @property
    def full_name(self) -> str:
        return self.com_obj.FullName

    @property
    def major(self) -> int:
        return self.com_obj.major

    @property
    def minor(self) -> int:
        return self.com_obj.minor

    @property
    def name(self) -> str:
        return self.com_obj.Name

    @property
    def patch(self) -> int:
        return self.com_obj.Patch
