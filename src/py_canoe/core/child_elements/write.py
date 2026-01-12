from typing import Union

from py_canoe.helpers.common import logger


class Write:
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def text(self) -> Union[str, None]:
        try:
            text_data: str = self.com_object.Text
            logger.info("📖 Text read successfully from write window")
            for line in text_data.splitlines():
                logger.info(f"    {line}")
            return text_data
        except Exception as e:
            logger.error(f"❌ Error getting text from write window: {e}")
            return None

    def clear(self) -> bool:
        try:
            self.com_object.Clear()
            logger.info("🧹 Write window cleared successfully")
            return True
        except Exception as e:
            logger.error(f"❌ Error clearing write window: {e}")
            return False

    def copy(self) -> bool:
        try:
            self.com_object.Copy()
            logger.info("📷 Write Window content copied to clipboard successfully")
            return True
        except Exception as e:
            logger.error(f"❌ Error copying write window content: {e}")
            return False

    def enable_output_file(self, output_file: str, tab_index=None) -> bool:
        try:
            if tab_index is not None:
                self.com_object.EnableOutputFile(output_file, tab_index)
            else:
                self.com_object.EnableOutputFile(output_file)
            logger.info(f"✔️ Enabled write window output file: {output_file}")
            return True
        except Exception as e:
            logger.error(f"❌ Error enabling write window output file '{output_file}': {e}")
            return False

    def disable_output_file(self, tab_index=None) -> bool:
        try:
            if tab_index is not None:
                self.com_object.DisableOutputFile(tab_index)
            else:
                self.com_object.DisableOutputFile()
            logger.info("⏹️ Disabled write window output file")
            return True
        except Exception as e:
            logger.error(f"❌ Error disabling write window output file: {e}")
            return False

    def output(self, text: str) -> bool:
        try:
            self.com_object.Output(text)
            logger.info(f"✍️ Text written in write window: {text}")
            return True
        except Exception as e:
            logger.error(f"❌ Error writing text in write window: {e}")
            return False
