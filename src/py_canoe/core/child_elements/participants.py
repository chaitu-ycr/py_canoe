from py_canoe.core.child_elements.participant import Participant


class Participants:
    """
    Collection of Participant objects associated with an ApplicationModel.
    """
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> 'Participant':
        return Participant(self.com_object.Item(index))

    def add(self, participant_path: str) -> 'Participant':
        return Participant(self.com_object.Add(participant_path))

    def remove(self, index: int):
        self.com_object.Remove(index)
