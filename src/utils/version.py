# Import Python Libraries here
import win32com.client

class Version:
    """The Version object represents the version of the CANoe application.
    """
    def __init__(self, app_obj) -> None:
        self.app_obj = app_obj
        self.log = self.app_obj.log
        self.ver_obj = win32com.client.Dispatch(self.app_obj.app_com_obj.Version)
    
    @property
    def build(self) -> int:
        """Returns the build number of the CANoe application.

        Returns:
            int: The build number of the CANoe application.
        """
        return self.ver_obj.Build

    @property
    def full_name(self) -> str:
        """Determines the complete path of the object.

        Returns:
            str: The complete CANoe version in the following format: "Vector CANoe /run 6.0.50" or "Vector CANoe.LIN /run 6.0.50".
        """
        return self.ver_obj.FullName
    
    @full_name.setter
    def full_name(self, full_name: str) -> None:
        """Sets the complete path of the object.

        Args:
            full_name (str): The complete CANoe version in the following format: "Vector CANoe /run 6.0.50" or "Vector CANoe.LIN /run 6.0.50".
        """
        self.ver_obj.FullName = full_name
        self.log.info(f'CANoe version set to {full_name}.')

    @property
    def major(self) -> int:
        """Returns the major version number of the CANoe application.

        Returns:
            int: The major version number of the CANoe application.
        """
        return self.ver_obj.major

    @property
    def minor(self) -> int:
        """Returns the Minor version number of the CANoe application.

        Returns:
            int: The Minor version number of the CANoe application.
        """
        return self.ver_obj.minor

    @property
    def name(self) -> str:
        """Returns the name of the object.

        Returns:
            str: The CANoe version in the following format: "CANoe 5.1 SP2" (with Service Pack) or "CANoe.LIN 5.1" (without Service Pack).
        """
        return self.ver_obj.Name

    @property
    def patch(self) -> int:
        """Returns the patch number of the CANoe application.

        Returns:
            int: The patch number of the CANoe application.
        """
        return self.ver_obj.Patch
