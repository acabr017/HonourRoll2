import pandas as pd
import os
import csv

files = ["A Average Citizenship Honor Roll","Citizenship Honor Roll","Principal Honor Roll","Regular Honor Roll","Superior Honor Roll", "totals"]


def set_save_directory(save_dir:str):
    global save_directory
    save_directory = save_dir.replace('/','\\')


def set_filename(file_name:str):
    global filename
    filename = file_name


def file_clean_up(file_type:str):
    for item in files:
        item = item.replace(" ","").lower() + "_" + filename
        path = os.path.join(save_directory,item) + "." + file_type
        print(f"Using path.join: {path}")
        print(os.path.exists(path))
        if os.path.exists(path):
            os.remove(path)
    print("Cleaned up!")


def final_file_clean_up(file_type:str):
    to_delete = "_" + filename + "." + file_type
    path = os.path.join(save_directory,to_delete)
    if os.path.exists(path):
            os.remove(path)


def Excel_to_text(sheet_as_string:str):
    currentRoll = ""
    for index, row in enumerate(sheet_as_string):
        if index == 0:
            continue
        row = row[2:]
        if "roll" in row.lower():
            rollType = row[len(row)-32::1].strip().replace(" ","").lower()
            currentRoll = rollType
            continue
        if ("STU ID" in row) or ("NaN" in row):
            continue
        items = list(filter(lambda x: len(x) > 2, row.split("               ")))
        for val in items:
            with open(f"{save_directory+'/'+currentRoll+'_'+filename}.txt", "a") as f:
                print(val.strip(),file=f)


def Excel_to_csv(sheet_as_string:str):
    currentRoll = ""
    for index, row in enumerate(sheet_as_string):
        if index == 0:
            continue
        row = row[2:]
        if "roll" in row.lower():
            rollType = row[len(row)-32::1].strip().replace(" ","").lower()
            currentRoll = rollType
            continue
        if ("STU ID" in row) or ("NaN" in row):
            continue
        items = list(filter(lambda x: len(x) > 2, row.split("               ")))
        with open(f"{save_directory+'/'+currentRoll+'_'+filename}.csv", "a") as f:
            writer = csv.writer(f)
            with open(f"{save_directory+'/'+currentRoll+'_'+filename}.csv", "r") as csvfile:
                if sum(1 for line in csvfile) == 0:
                    header = ['STU ID',' GR-HR', 'STUDENT NAME']
                    writer.writerow(header)
            for line in items:
                line = ' '.join(line.split("  ")).split()
                writer.writerow(line)


def Line_counter():
    with open(f"{save_directory}/totals_{filename}.txt", "a") as f1:
        for item in files:
            fileName = item.replace(" ","").lower() + '_' + filename
            if fileName == "totals":
                continue
            try:
                with open(f"{fileName}.txt", 'r') as f2:
                    amountOfLines = len(f2.readlines())
                    print(f"{item}: {amountOfLines}", file=f1)
            except FileNotFoundError:
                print(f"{item}: 0", file=f1)
        

def create_dataframes(filepath:str):
    return pd.read_excel(filepath, skiprows=7, header=None, sheet_name=None)


def convert_dataframe(dataframe,output_type:str):
    for key, values in dataframe.items():
        if key == "Sheet1" or key == f"Sheet{len(dataframe.values())}":
            continue
        excel_string = values.to_string().split("\n")
        if output_type.lower() == 'txt':
            Excel_to_text(excel_string)
        elif output_type.lower() == 'csv':
            Excel_to_csv(excel_string)




# test things

def testing_globals():
    print(f"Global Filename: {filename} and Global Directory: {save_directory}")


def testing_file_creation():
    for item in files:
        item = item.replace(" ","").lower() + " " + filename
        with open(item,"w") as f:
            print("Testing", file=f)