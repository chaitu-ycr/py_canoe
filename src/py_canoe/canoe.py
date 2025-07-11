import logging

from py_canoe.utils.application import Application

logging.getLogger('py_canoe')

class CANoe:
    def __init__(self, py_canoe_log_dir='', user_capl_functions=tuple()):
        self.application = Application()

    def new(self, auto_save: bool = False, prompt_user: bool = False):
        try:
            self.application.new(auto_save, prompt_user)
            logging.info('📢 New CANoe configuration successfully created 🎉')
        except Exception as e:
            logging.error(f"😡 Error creating new CANoe configuration: {e}")

    def open(self, canoe_cfg: str, visible=True, auto_save=True, prompt_user=False, auto_stop=True) -> None:
        try:
            self.application.open(canoe_cfg, auto_save, prompt_user)
            self.application.visible = visible
            logging.info(f'📢 Opened CANoe configuration: {canoe_cfg} 🎉')
        except Exception as e:
            logging.error(f"😡 Error opening CANoe configuration '{canoe_cfg}': {e}")

    def quit(self):
        try:
            self.application.quit()
            logging.info('📢 CANoe application quit successfully 🎉')
        except Exception as e:
            logging.error(f"😡 Error quitting CANoe application: {e}")
