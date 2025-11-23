class CompileResult:
    """
    The CompileResult object represents the result of the last compilation of the CAPL object.
    """
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def error_message(self) -> str:
        return self.com_object.ErrorMessage

    @property
    def node_name(self) -> str:
        return self.com_object.NodeName

    @property
    def result(self) -> int:
        return self.com_object.result

    @property
    def source_file(self) -> str:
        return self.com_object.SourceFile
