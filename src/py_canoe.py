# import external modules here

# import internal modules here
from py_canoe_app.py_canoe_logger import PyCanoeLogger
from py_canoe_app.application import Application


class CANoe:
    def __init__(self, py_canoe_log_dir='', user_capl_functions=tuple()):
        self.log = PyCanoeLogger(py_canoe_log_dir).log
        self.user_capl_function_names = user_capl_functions
        self.application = Application()


# canoe_inst = CANoe()

# print('canoe_inst.application.system'.center(120, '-'))
# print(canoe_inst.application.system.namespaces_count)
# canoe_inst.application.system.add_system_variable('dummy', 'var', 321)
# canoe_inst.application.system.remove_system_variable('dummy::group', 'var1')
# canoe_inst.application.system.remove_system_variable('hello', 'var2')

# print('canoe_inst.application.ui'.center(120, '-'))
# print(canoe_inst.application.ui.clear_write_window())
# print(canoe_inst.application.ui.get_write_window_text)

# print('canoe_inst.application.version'.center(120, '-'))
# print(canoe_inst.application.version.build)
# print(canoe_inst.application.version.full_name)
# print(canoe_inst.application.version.major)
# print(canoe_inst.application.version.minor)

# print('Hello End')