from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from py_canoe.canoe import CANoe


class VTest:
    """VTest class for controlling CANoe test modules"""
    def __init__(self, canoe_inst: "CANoe"):
        self.canoe = canoe_inst
        self.canoe.application.configuration.fetch_test_modules()

    def list_test_environments(self) -> list:
        """returns list of test environments in loaded canoe configuration.

        Returns:
            list: list of test environments.
        """
        return list(self.canoe.get_test_environments().keys())

    def list_test_modules(self, env_name: str) -> list:
        """returns list of test modules in a given test environment.

        Args:
            env_name (str): test environment name.

        Returns:
            list: list of test modules in given test environment.
        """
        return list(self.canoe.get_test_modules(env_name).keys())

    def run_test_module(self, module_name: str) -> int:
        """runs a single test module.

        Args:
            module_name (str): test module name.

        Returns:
            int: test module execution verdict.
        """
        return self.canoe.execute_test_module(module_name)

    def run_test_environment(self, env_name: str) -> None:
        """runs all test modules in a given test environment.

        Args:
            env_name (str): test environment name.
        """
        self.canoe.execute_all_test_modules_in_test_env(env_name)

    def run_all_tests(self) -> None:
        """runs all test modules in loaded canoe configuration."""
        self.canoe.execute_all_test_environments()
