import win32com.client


class SecuritySetup:
    def __init__(self, security_setup_com_object):
        self.com_object = win32com.client.Dispatch(security_setup_com_object)

    @property
    def tls_observer_security_configuration(self) -> 'SecurityConfiguration':
        return SecurityConfiguration(self.com_object.TLSObserverSecurityConfiguration)


class SecurityConfiguration:
    def __init__(self, security_configuration_com_object):
        self.com_object = win32com.client.Dispatch(security_configuration_com_object)

    @property
    def security_active(self) -> bool:
        return self.com_object.SecurityActive

    @security_active.setter
    def security_active(self, value: bool):
        self.com_object.SecurityActive = value

    @property
    def security_profile(self) -> str:
        return self.com_object.SecurityProfile

    @security_profile.setter
    def security_profile(self, value: str):
        self.com_object.SecurityProfile = value
