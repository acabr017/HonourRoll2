import pandas as pd
import os
import csv


class ExcelConvert:
    """

    Args: output_path - Where the post-converted files will be created
          file_name - the name of the file, without file extensions

    """

    FILES = [
        "A Average Citizenship Honor Roll",
        "Citizenship Honor Roll",
        "Principal Honor Roll",
        "Regular Honor Roll",
        "Superior Honor Roll",
        "totals",
    ]

    def __init__(self, output_path: str, file_name: str, output_type: str):
        self.original_filename = file_name
        self.output_path = output_path
        self.output_type = output_type

    def create_dataframes(self, filepath: str):
        return pd.read_excel(filepath, skiprows=7, header=None, sheet_name=None)

    def convert_dataframe(self, dataframe, output_type: str):
        for key, values in dataframe.items():
            if key == "Sheet1" or key == f"Sheet{len(dataframe.values())}":
                continue
            excel_string = values.to_string().split("\n")
            if output_type.lower() == "txt":
                self.Excel_to_text(excel_string)
            elif output_type.lower() == "csv":
                self.Excel_to_csv(excel_string)

    def Excel_to_text(self, sheet_as_string_list: list):
        currentRoll = ""
        warning_page = False
        # If this is true, this page is a warning page. Do not record
        for index, row in enumerate(sheet_as_string_list):
            if index != 0 and (
                ("STU ID" not in row) and ("NaN" not in row)
            ):  # Skipping because the first element is blank for some reason.
                row = row[2:]
                if "roll" in row.lower():
                    rollType = row[len(row) - 32 :: 1].strip().replace(" ", "").lower()
                    currentRoll = rollType
                    continue
                if "WARNING MESSAGES" in row:
                    warning_page = True
                items = list(filter(lambda x: len(x) > 2, row.split("               ")))
                if warning_page is False:
                    for val in items:
                        with open(
                            f"{self.output_path+'/'+currentRoll+'_'+self.original_filename}.txt",
                            "a",
                        ) as f:
                            print(val.strip(), file=f)

    def Excel_to_csv(self, sheet_as_string: str):
        currentRoll = ""
        warning_page = False
        # If this is true, this page is a warning page. Do not record
        for index, row in enumerate(sheet_as_string):
            if index != 0 and (
                ("STU ID" not in row) and ("NaN" not in row)
            ):  # Skipping because the first element is blank for some reason.
                row = row[2:]
                if "roll" in row.lower():
                    rollType = row[len(row) - 32 :: 1].strip().replace(" ", "").lower()
                    currentRoll = rollType
                    continue
                if "WARNING MESSAGES" in row:
                    warning_page = True
                items = list(filter(lambda x: len(x) > 2, row.split("               ")))
                if warning_page is False:
                    filepath = f"{self.output_path+'/'+currentRoll+'_'+self.original_filename}.csv"
                    with open(filepath, "a+", newline="") as f:
                        writer = csv.writer(f)
                        for line in items:
                            line = " ".join(line.split("  ")).split()
                            writer.writerow(line)

    def create_output_files(self):
        for type in self.FILES:
            if type == "totals":
                continue
            filepath = (
                self.output_path
                + "/"
                + type.replace(" ", "").lower()
                + "_"
                + self.original_filename
                + "."
                + self.output_type
            )
            with open(filepath, "a", newline="") as setup_file:
                if self.output_type == "csv":
                    writer = csv.writer(setup_file)
                    header = ["STU ID", " GR-HR", "STUDENT NAME"]
                    writer.writerow(header)
                elif self.output_type == "txt":
                    print("STU ID  GR-HR   STUDENT NAME", file=setup_file)

    def Line_counter(self):
        with open(f"{self.output_path}/totals_{self.original_filename}.txt", "a") as f1:
            for item in self.FILES:
                fileName = item.replace(" ", "").lower() + "_" + self.original_filename
                if "totals" in fileName:
                    continue
                try:
                    with open(
                        f"{self.output_path}/{fileName}.{self.output_type}", "r"
                    ) as f2:
                        amountOfLines = (
                            len(f2.readlines()) - 1
                        )  # Subtract 1 for the header
                        print(f"{item}: {amountOfLines}", file=f1)
                except FileNotFoundError:
                    print(f"{item}: 0", file=f1)

    def final_file_clean_up(self, file_type: str):
        to_delete = "_" + self.original_filename + "." + file_type
        path = os.path.join(self.output_path, to_delete)
        if os.path.exists(path):
            os.remove(path)

    def do_conversion(self, input_path, output_type):
        df = self.create_dataframes(input_path)
        self.convert_dataframe(df, output_type)
        self.Line_counter()
