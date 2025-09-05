from py_canoe.core.child_elements.test_module import TestModule


class TestModules:
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def count(self) -> int:
        return self.com_object.Count

    def add(self, full_name: str) -> 'TestModule':
        return TestModule(self.com_object.Add(full_name))

    def remove(self, index: int, prompt_user=False) -> None:
        self.com_object.Remove(index, prompt_user)

    def fetch_test_modules(self) -> dict['str': 'TestModule']:
        test_modules = dict()
        for index in range(1, self.count + 1):
            tm_inst = TestModule(self.com_object.Item(index))
            test_modules[tm_inst.name] = tm_inst
        return test_modules
