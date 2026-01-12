import win32com.client

from py_canoe.core.child_elements.test_environments import TestEnvironments


class TestSetup:
    """The TestSetup object represents CANoe's test setup."""
    def __init__(self, com_object):
        self.com_object = win32com.client.Dispatch(com_object)

    def save_all(self, prompt_user=False) -> None:
        self.com_object.SaveAll(prompt_user)

    @property
    def test_environments(self) -> 'TestEnvironments':
        return TestEnvironments(self.com_object.TestEnvironments)
