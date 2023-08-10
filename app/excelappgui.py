import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import excel_converter as ec
import os
import errno
from classes import Excel_app_helper
from pathlib import Path
import shutil


files = Excel_app_helper()
get_files_path = "C://"


def start_button_state():
    if len(files.input_files) > 0 and files.out_location_bool and files.out_type_bool:
        start_button["state"] = tk.NORMAL
        start_button["fg"] = "green"
    else:
        start_button["state"] = tk.DISABLED
        start_button["bg"] = "white"


def get_files_func():
    files.input_files = list(
        filedialog.askopenfilenames(initialdir=get_files_path, title="Select File(s)")
    )
    files.set_files_dict()
    for file_name, path in files.files_dict.items():
        if file_name.endswith(".xls") or file_name.endswith(".xlsx"):
            insert_to_table("input", file_name, path)
        else:
            messagebox.showerror(
                "Warning!", "Only .xls and .xlsx files accepted!", parent=window
            )
    start_button_state()


def set_directory_func():
    files.set_output_location(filedialog.askdirectory() + "/Output Files")
    directory_label_string.set(files.output_location)
    print(files.output_location)
    start_button_state()


def radio_button_func():
    files.set_output_type(output_type_string.get())
    start_button_state()


def insert_to_table(table_type: str, f_name: str, f_path: str):
    if table_type.lower() == "input":
        input_table.insert(parent="", index=tk.END, values=(f_name, f_path))
    elif table_type.lower() == "output":
        output_table.insert(parent="", index=tk.END, values=(f_name, f_path))
    else:
        print("Invalid table type!")


def start_button_func():
    dirpath = Path(files.output_location)
    if dirpath.exists() and dirpath.is_dir():
        shutil.rmtree(dirpath)
    try:
        make_new_dir(files.output_location)
    except OSError as exc:
        if exc.errno != errno.EEXIST:
            raise
        pass
    for file_name, path in files.files_dict.items():
        converter = ec.ExcelConvert(
            files.output_location, file_name.split(".")[0], files.output_type
        )
        converter.create_output_files()
        converter.do_conversion(path, files.output_type)

    populate_output_table()


def make_new_dir(dir_path: str):
    os.makedirs(dir_path)


def get_output_files():
    return [
        fn
        for fn in os.listdir(files.output_location)
        if (fn.lower().endswith(files.output_type) or fn.lower().endswith("txt"))
        and os.path.isfile(os.path.join(files.output_location, fn))
    ]


def populate_output_table():
    output_files = get_output_files()

    print(output_files)
    for file in output_files:
        insert_to_table(
            "output",
            file,
            f"{files.output_location}/{file}",
        )


def toggle_delete(_):
    print(f"Event Triggered - argument is {_}")
    if len(input_table.selection()) > 0:
        file_delete_button["state"] = tk.NORMAL
        file_delete_button["fg"] = "red"
    else:
        file_delete_button["state"] = tk.DISABLED
        file_delete_button["bg"] = "white"
    start_button_state()


def delete_items(_):
    print(f"Event Triggered - argument is {_}")
    for i in input_table.selection():
        files.input_files.remove(input_table.item(i)["values"][1])
        input_table.delete(i)


# -------------- Window --------------


window = tk.Tk()


window.title("Val's App")
window.geometry("800x900")
window.resizable(False, False)

window_width = int(window.winfo_screenmmwidth())
window_height = int(window.winfo_screenmmheight())

# -------------- Widgets --------------

# :::::::::::::::::: Frames ::::::::::::::::::


# Instructions Frame
instructions_frame = tk.Frame(
    window, borderwidth=5, width=570, height=150, relief="groove"
)
# instructions_frame.pack_propagate(False)

# Setup Buttons Frame
setup_buttons_frame = tk.Frame(
    window, borderwidth=5, width=int(window_width * 1.4), height=100
)
setup_buttons_frame.pack_propagate(False)

# Output Files Label Frame
output_files_label_frame = tk.Frame(
    window, width=int(window_width * 1.2), height=50, relief="groove"
)
setup_buttons_frame.pack_propagate(False)

# Frame to hold table label Frame
placeholder_frame = tk.Frame(window, width=int(window_width * 1.2), height=50)
placeholder_frame.pack_propagate(False)

# Table Label Frame
table_label_frame = tk.Frame(
    placeholder_frame, width=int(window_width * 1.2), height=40, relief="groove"
)
table_label_frame.pack_propagate(False)

# Frame for Radio Buttons
radio_frame = ttk.Frame(setup_buttons_frame, width=200, height=80)
radio_frame.pack_propagate(False)

# Frame for Start and Delete Buttons
del_and_start_buttons_frame = tk.Frame(
    window, width=int(window_width * 2), height=50, relief="groove"
)
del_and_start_buttons_frame.pack_propagate(False)

# Input Table Frame
input_table_frame = tk.Frame(window)

# :::::::::::::::::: Buttons ::::::::::::::::::

# Get File Button
get_files_button = tk.Button(
    setup_buttons_frame,
    text="Upload Files",
    command=get_files_func,
    font=("Times New Roman", 14),
)

# File Delete Button
file_delete_button = tk.Button(
    del_and_start_buttons_frame,
    text="                   \
                                Delete                   ",
    state="disabled",
    command=lambda: delete_items(""),
    font=("Times New Roman", 14),
)

# Set Directory Button
set_directory_button = tk.Button(
    setup_buttons_frame,
    text="Set Directory For Output Files",
    command=set_directory_func,
    font=("Times New Roman", 14),
)

# Start Button
start_button = tk.Button(
    del_and_start_buttons_frame,
    text="                    \
                            Start                    ",
    state="disabled",
    command=start_button_func,
    font=("Times New Roman", 14),
)

# Radio Buttons for Output Type:
output_type_string = tk.StringVar()


txt_radio_button = ttk.Radiobutton(
    radio_frame,
    text="Text Output",
    value="txt",
    variable=output_type_string,
    command=radio_button_func,
)
csv_radio_button = ttk.Radiobutton(
    radio_frame,
    text="CSV Output",
    value="csv",
    variable=output_type_string,
    command=radio_button_func,
)


# :::::::::::::::::: Treeview :::::::::::::::::
# Input and Output Tables

style = ttk.Style()
style.configure("TRadiobutton", font="timesnewroman 12", foreground="black")

input_table = ttk.Treeview(input_table_frame, columns=(1, 2), show="headings")

input_table.heading(1, text="Input File Name")
input_table.heading(2, text="Input File Path")
input_table.column(1, width=15)


output_table = ttk.Treeview(window, columns=("file", "filepath"), show="headings")
output_table.heading("file", text="Output File Name")
output_table.heading("filepath", text="Output File Path")
output_table.column("file", width=15)


# :::::::::::::::::: Labels ::::::::::::::::::

# Label for Directory
directory_label_string = tk.StringVar()
output_labels_frame_label = tk.Label(
    output_files_label_frame, text="Files will be saved to:"
)
directory_label = tk.Label(
    output_files_label_frame, textvariable=directory_label_string
)

# Instructions Label
instructions_text1 = "1. Upload the files to be converted.\n\
2. Specify the output location.\n\
3. Specify output file type (txt/csv)"

instructions_text2 = 'Files will be generated in specified output location, \
labled as "(rolltype)_(filename).(outputtype)"\n \
Example: Uploaoding file: 5052 Q1.xls -> "regularhonorroll_5052Q1.txt"'


instructions_label1 = tk.Label(
    instructions_frame,
    justify="left",
    text=instructions_text1,
    font=("Times New Roman", 14),
)
instructions_label2 = tk.Label(
    instructions_frame,
    justify="center",
    text=instructions_text2,
    font=("Times New Roman", 12),
)


# Table Label

table_label = tk.Label(
    table_label_frame,
    justify="left",
    text="Selected Files",
    font=("Times New Roman", 11),
)

# :::::::::::::::::: Events ::::::::::::::::::
# events
input_table.bind("<<TreeviewSelect>>", toggle_delete)
# table.bind('<Delete>', delete_items)


#  -------------- Packing --------------
# 1st Frame - Instructions
instructions_frame.pack()
instructions_label1.pack()
instructions_label2.pack(side="bottom")

# 2nd Frame - Upload button, directory button, output type radio buttons
setup_buttons_frame.pack(expand=True)
get_files_button.pack(side="left")
set_directory_button.pack(side="left", padx=10)
radio_frame.pack(pady=20, side="left")
txt_radio_button.pack()
csv_radio_button.pack()

# 3rd Frame - Output location labels
output_files_label_frame.pack(expand=True)
output_labels_frame_label.pack(side="left")
directory_label.pack(side="right")

# 4th Frame - Table Label
placeholder_frame.pack()
table_label_frame.pack(side="left")
table_label.pack(side="left")

# 5th Frame - Input Table
input_table_frame.pack(fill="both")
input_table.pack(padx=20, fill="x")

# 6th Frame
del_and_start_buttons_frame.pack()
start_button.pack(side="left", expand=True)
file_delete_button.pack(expand=True, side="right")

# Output Table
output_table.pack(padx=20, fill="x")


# -------------- Run --------------

window.mainloop()
