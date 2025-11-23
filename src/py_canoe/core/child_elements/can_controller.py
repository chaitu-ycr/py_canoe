import win32com.client


class CanController:
    def __init__(self, com_object):
        self.com_object = win32com.client.Dispatch(com_object)

    @property
    def acknowledge(self):
        return self.com_object.Acknowledge

    @acknowledge.setter
    def acknowledge(self, value: bool):
        self.com_object.Acknowledge = value

    @property
    def baudrate(self):
        return self.com_object.Baudrate

    @baudrate.setter
    def baudrate(self, value: int):
        self.com_object.Baudrate = value

    @property
    def btr0(self):
        return self.com_object.BTR0

    @property
    def btr1(self):
        return self.com_object.BTR1

    @property
    def output_control(self):
        return self.com_object.OutputControl

    @output_control.setter
    def output_control(self, value: int):
        self.com_object.OutputControl = value

    @property
    def samples(self):
        return self.com_object.Samples

    @samples.setter
    def samples(self, value: int):
        self.com_object.Samples = value

    @property
    def self_ack_enabled(self):
        return self.com_object.SelfAckEnabled

    @self_ack_enabled.setter
    def self_ack_enabled(self, value: bool):
        self.com_object.SelfAckEnabled = value

    @property
    def synchronisation(self):
        return self.com_object.Synchronisation

    @synchronisation.setter
    def synchronisation(self, value: int):
        self.com_object.Synchronisation = value

    def can_set_config(self, baudrate: int, tseg1: int, tseg2: int, sjw: int, sam: int, flags: int):
        self.com_object.CANSetConfig(baudrate, tseg1, tseg2, sjw, sam, flags)

    def can_set_fd_arb_phase_config(self, baudrate: int, tseg1: int, tseg2: int, sjw: int, sam: int, flags: int):
        self.com_object.CANSetFDArbPhaseConfig(baudrate, tseg1, tseg2, sjw, sam, flags)

    def can_set_fd_data_phase_config(self, baudrate: int, tseg1: int, tseg2: int, sjw: int, sam: int, flags: int):
        self.com_object.CANSetFDDataPhaseConfig(baudrate, tseg1, tseg2, sjw, sam, flags)

    def set_btr(self, btr0: int, btr1: int):
        self.com_object.SetBTR(btr0, btr1)
