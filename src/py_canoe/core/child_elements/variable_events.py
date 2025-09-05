class VariableEvents:
    def __init__(self) -> None:
        self.VARIABLE_INFO = {}
        self.VARIABLE_UPDATED = False

    def OnChange(self, value):
        self.VARIABLE_INFO['value'] = value
        self.VARIABLE_UPDATED = True

    def OnChangeAndTime(self, value, timeHigh, time):
        self.VARIABLE_INFO['value'] = value
        self.VARIABLE_INFO['timeHigh'] = timeHigh
        self.VARIABLE_INFO['time'] = time
        self.VARIABLE_UPDATED = True
