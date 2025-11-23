class MostDisassembler:
    def __init__(self, com_object):
        self.com_object = com_object

    def symbolic_message_id_components(self, f_block_id: int, instance_id: int, function_id: int, op_type_id: int) -> int:
        return self.com_object.SymbolicMessageIDComponents(f_block_id, instance_id, function_id, op_type_id)

    def symbolic_parameter_list1(self, data_length: int, data_array: bytearray, max_params: int = 0) -> tuple:
        return self.com_object.SymbolicParameterList1(data_length, data_array, max_params)

    def symbolic_parameter_list2(self, f_block_id: int, instance_id: int, function_id: int, op_type_id: int, data_length: int, data_array: bytearray, max_params: int = 0) -> tuple:
        return self.com_object.SymbolicParameterList2(f_block_id, instance_id, function_id, op_type_id, data_length, data_array, max_params)

    def this_message_id_components(self, f_block_id: int, instance_id: int, function_id: int, op_type_id: int) -> int:
        return self.com_object.ThisMessageIDComponents(f_block_id, instance_id, function_id, op_type_id)

    def this_symbolic_message_id_components(self, f_block_name: str, function_name: str, op_type_name: str) -> int:
        return self.com_object.ThisSymbolicMessageIDComponents(f_block_name, function_name, op_type_name)

    def this_symbolic_parameter_list(self, max_params: int = 0) -> tuple:
        return self.com_object.ThisSymbolicParameterList(max_params)
