import pandas as pd
import os
import csv


class ExcelUtility:
    FILES = [
        "A Average Citizenship Honor Roll",
        "Citizenship Honor Roll",
        "Principal Honor Roll",
        "Regular Honor Roll",
        "Superior Honor Roll",
        "totals",
    ]

    ROW_OFFSET = 32

    DEFAULT_HEADER = ["STU ID", " GR-HR", "STUDENT NAME"]

    def __init__(self, save_directory: str = "", filename: str = ""):
        self.save_directory = save_directory.replace("/", "\\")
        self.filename = filename

    def file_clean_up(self, file_type: str):
        for file_descriptor in self.FILES:
            file = file_descriptor.replace(" ", "").lower() + "_" + self.filename
            path = os.path.join(self.save_directory, file) + "." + file_type

            if os.environ.get("ENABLE_DEBUG", False):
                print(f"Using path.join: {path}")
                print(os.path.exists(path))

            if os.path.exists(path):
                os.remove(path)

        if os.environ.get("ENABLE_DEBUG", False):
            print("Cleaned up!")

    def final_file_clean_up(self, file_type: str):
        to_delete = "_" + self.filename + "." + file_type
        path = os.path.join(self.save_directory, to_delete)

        if os.path.exists(path):
            os.remove(path)

    def excel_to_text(self, sheet_as_string: str):
        current_roll = ""

        _headers = sheet_as_string[0]
        data = sheet_as_string[1:]
        for index, row in enumerate(data):
            row = row[2:]

            if "roll" in row.lower():
                roll_type = (
                    row[len(row) - self.ROW_OFFSET :: 1]
                    .strip()
                    .replace(" ", "")
                    .lower()
                )
                current_roll = roll_type

            if ("STU ID" in row) or ("NaN" in row):
                continue

            items = list(filter(lambda x: len(x) > 2, row.split("               ")))

            for val in items:
                with open(
                    f"{self.save_directory + '/' + current_roll + '_ '+ self.filename}.txt",
                    "a",
                ) as f:
                    f.write(val.strip())

    def excel_to_csv(self, sheet_as_string: str):
        current_roll = ""

        _headers = sheet_as_string[0]
        data = sheet_as_string[1:]
        for index, row in enumerate(data):
            row = row[2:]

            if "roll" in row.lower():
                roll_type = (
                    row[len(row) - self.ROW_OFFSET :: 1]
                    .strip()
                    .replace(" ", "")
                    .lower()
                )
                current_roll = roll_type

            if ("STU ID" in row) or ("NaN" in row):
                continue

            items = list(filter(lambda x: len(x) > 2, row.split("               ")))

            with open(
                f"{self.save_directory + '/' + current_roll + '_' + self.filename}.csv",
                "a",
            ) as f:
                writer = csv.writer(f)

                with open(
                    f"{self.save_directory + '/' + current_roll + '_' + self.filename}.csv",
                    "r",
                ) as csvfile:
                    if sum(1 for line in csvfile) == 0:
                        writer.writerow(self.DEFAULT_HEADER)

                for line in items:
                    line = " ".join(line.split("  ")).split()
                    writer.writerow(line)

    def line_count(self):
        with open(f"{self.save_directory}/totals_{self.filename}.txt", "a") as f1:
            for item in self.FILES:
                file_name = item.replace(" ", "").lower() + "_" + self.filename
                if file_name == "totals":
                    continue
                try:
                    with open(f"{file_name}.txt", "r") as f2:
                        line_count = len(f2.readlines())
                        f1.write(f"{item}: {line_count}")
                except FileNotFoundError:
                    f1.write(f"{item}: 0")

    def create_dataframes(self, filepath: str):
        return pd.read_excel(filepath, header=None, sheet_name=None)

    def convert_dataframe(self, dataframe, output_type: str):
        for key, values in dataframe.items():
            if key == "Sheet1" or key == f"Sheet{len(dataframe.values())}":
                continue

            excel_string = values.to_string().split("\n")

            if output_type.lower() == "txt":
                self.excel_to_text(excel_string)
            elif output_type.lower() == "csv":
                self.excel_to_csv(excel_string)
