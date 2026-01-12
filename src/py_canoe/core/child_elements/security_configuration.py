class SecurityConfiguration:
    """The SecurityConfiguration object represents a security profile assignment to a network, TCP stack or observer."""
    def __init__(self, security_configuration_com_obj):
        self.com_object = security_configuration_com_obj

    @property
    def security_active(self) -> bool:
        return self.com_object.SecurityActive

    @security_active.setter
    def security_active(self, value: bool):
        self.com_object.SecurityActive = value

    @property
    def security_profile(self) -> int:
        return self.com_object.SecurityProfile

    @security_profile.setter
    def security_profile(self, value: int):
        self.com_object.SecurityProfile = value
