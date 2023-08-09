class Excel_app_helper:
    def __init__(self):
        self.input_files = []
        self.files_dict = {}
        self.output_location = ""
        self.output_type = ""
        self.out_location_bool = False
        self.out_type_bool = False

    def set_output_type(self, type: str):
        self.output_type = type
        self.out_type_bool = True if len(type) > 0 else False

    def set_output_location(self, path: str):
        self.output_location = path
        self.out_location_bool = True if len(path) > 0 else False

    def set_files_dict(self):
        if len(self.input_files) == 0:
            # TO DO: Raise an error if this is empty
            print("this is empty")
            return

        for val in self.input_files:
            file_name = val.split("/")[-1].replace(" ", "")
            self.files_dict[file_name] = val


# def testing_output(obj):
#     print(
#         f"\t\tTesting:\n \
#     Input: {obj.input_files} \n \
#     Input Dictionary: {obj.files_dict} \n \
#     Output_location: {obj.output_location} \n \
#     Output_type: {obj.output_type} \n \
#     Output_bool: {obj.out_location_bool} \n \
#     Type_bool: {obj.out_type_bool}"
#     )


# test = Excel_app_helper()
# testing_output(test)

# test.set_output_type("txt")
# testing_output(test)


# sample_paths = [
#     "C://user/ariel/data1.xls",
#     "C://user/ariel/data2.xls",
#     "C://user/ariel/data3.xls",
#     "C://user/ariel/data4.xls",
#     "C://user/ariel/data5.xls",
#     "C://user/ariel/data6.xls",
# ]
# test.input_files = sample_paths

# testing_output(test)
# test.set_files_dict()
# testing_output(test)
