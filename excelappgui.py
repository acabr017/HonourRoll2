import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import excel_converter as ec
import os
import errno

file_name = ""
output_location = ""
output_type = ""
output_locatin_selected = False
output_type_selected = False
get_files_path = "C://"


def start_button_state():
    if len(file_name) > 0 and output_locatin_selected and output_type_selected:
        start_button['state'] = tk.NORMAL
        start_button['fg'] = 'green'
    else: 
        start_button['state'] = tk.DISABLED
        start_button['bg'] = 'white'


def get_files_func():
    global file_name
    file_name = list(filedialog.askopenfilenames(initialdir=get_files_path,
                                                 title="Select File(s)"))
    for file in file_name:
        if file.endswith(".xls") or file.endswith(".xlsx"):
            insert_to_input_table(clean_string_for_table_insert(file), file)
        else:
            messagebox.showerror("Warning!",
                                 "Only .xls and .xlsx files accepted!",
                                 parent=window)
    start_button_state()


def clean_string_for_table_insert(string_to_clean):
    return string_to_clean.split("/")[-1].replace(" ", "")


def insert_to_input_table(f_name: str, f_path: str):
    input_table.insert(parent="", index=tk.END, values=(f_name, f_path))


def insert_to_output_table(f_name: str, f_path: str):
    output_table.insert(parent="", index=tk.END, values=(f_name, f_path))


def toggle_delete(_):
    print(f"Event Triggered - Condition is {len(file_name) > 0}")
    if len(input_table.selection()) > 0:
        file_delete_button['state'] = tk.NORMAL
        file_delete_button['fg'] = 'red'
    else: 
        file_delete_button['state'] = tk.DISABLED
        file_delete_button['bg'] = 'white'
    start_button_state()


def delete_items(_):
    for i in input_table.selection():
        file_name.remove(input_table.item(i)['values'][1])
        input_table.delete(i)


def set_directory_func():
    global output_location, output_locatin_selected, directory_label_string
    output_location = filedialog.askdirectory()
    output_locatin_selected = True
    output_location += "/Output Files"
    directory_label_string.set(output_location)
    start_button_state()
    ec.set_save_directory(output_location)


def get_filename(file_path_string: str):
    return file_path_string.split("/")[-1].replace(" ", "").split('.')[0]


def make_new_dir(dir_path: str):
    os.makedirs(dir_path)


def get_output_files():
    return [fn for fn in os.listdir(output_location) if
            fn.lower().endswith(output_type) and
            os.path.isfile(os.path.join(output_location, fn))]


def populate_output_table():
    output_files = get_output_files()
    print(output_files)
    for file in output_files:
        insert_to_output_table(clean_string_for_table_insert(file),
                               f"{output_location}/{file}")
    

def start_button_func():
    try:
        make_new_dir(output_location)
    except OSError as exc:
        if exc.errno != errno.EEXIST:
            raise
        pass
    for file in file_name:
        ec.set_filename(get_filename(file))
        ec.file_clean_up(output_type)
        ec.convert_dataframe(ec.create_dataframes(file), output_type)
        ec.Line_counter()
        ec.final_file_clean_up(output_type)
    populate_output_table()


def radio_button_func():
    global output_type_selected, output_type
    output_type_selected = True
    output_type = output_type_string.get()
    start_button_state()

# -------------- Window --------------


window = tk.Tk()


window.title("Val's App")
window.geometry("1000x900")
window.resizable(False, False)

window_width = int(window.winfo_screenmmwidth())
window_height = int(window.winfo_screenmmheight())

# -------------- Widgets --------------

# :::::::::::::::::: Frames :::::::::::::::::: 


# Instructions Frame
instructions_frame = tk.Frame(window,
                              borderwidth=5,
                              width=570,
                              height=150,
                              relief='groove')
# instructions_frame.pack_propagate(False)

# Setup Buttons Frame
setup_buttons_frame = tk.Frame(window,
                               borderwidth=5,
                               width=int(window_width*1.2),
                               height=100)
setup_buttons_frame.pack_propagate(False)

# Output Files Label Frame
output_files_label_frame = tk.Frame(window,
                                    width=int(window_width*1.2),
                                    height=50,
                                    relief='groove')
setup_buttons_frame.pack_propagate(False)

# Frame to hold table label Frame
placeholder_frame = tk.Frame(window,
                             width=int(window_width*1.2),
                             height=50)
placeholder_frame.pack_propagate(False)

# Table Label Frame
table_label_frame = tk.Frame(placeholder_frame,
                             width=int(window_width*1.2),
                             height=40, 
                             relief='groove')

table_label_frame.pack_propagate(False)

# Frame for Radio Buttons
radio_frame = ttk.Frame(setup_buttons_frame,
                        width=200,
                        height=80)
radio_frame.pack_propagate(False)

# Frame for Start and Delete Buttons
del_and_start_buttons_frame = tk.Frame(window,
                                       width=570,
                                       height=50)
del_and_start_buttons_frame.pack_propagate(False)

# Input Table Frame
input_table_frame = tk.Frame(window)

# :::::::::::::::::: Buttons :::::::::::::::::: 

# Get File Button
get_files_button = tk.Button(setup_buttons_frame,
                             text="Upload Files",
                             command=get_files_func,
                             font=('Times New Roman', 14))

# File Delete Button
file_delete_button = tk.Button(del_and_start_buttons_frame,
                               text="                   \
                                Delete                   ",
                               state="disabled",
                               command=lambda: delete_items(""),
                               font=('Times New Roman', 14))

# Set Directory Button
set_directory_button = tk.Button(setup_buttons_frame,
                                 text="Set Directory For Output Files",
                                 command=set_directory_func,
                                 font=('Times New Roman', 14))

# Start Button
start_button = tk.Button(del_and_start_buttons_frame,
                         text="                    \
                            Start                    ",
                         state='disabled',
                         command=start_button_func,
                         font=('Times New Roman', 14))

# Radio Buttons for Output Type:
output_type_string = tk.StringVar()


txt_radio_button = ttk.Radiobutton(radio_frame,
                                   text="Text Output",
                                   value="txt", 
                                   variable=output_type_string,
                                   command=radio_button_func)
csv_radio_button = ttk.Radiobutton(radio_frame,
                                   text="CSV Output",
                                   value='csv',
                                   variable=output_type_string,
                                   command=radio_button_func)


# :::::::::::::::::: Treeview :::::::::::::::::: 
# Input and Output Tables

style = ttk.Style()
style.configure('TRadiobutton',
                font='timesnewroman 12',
                foreground='black')

input_table = ttk.Treeview(input_table_frame,
                           columns=(1, 2),
                           show='headings')

input_table.heading(1,
                    text="Input File Name")
input_table.heading(2,
                    text="Input File Path")
input_table.column(1,
                   width=15)


output_table = ttk.Treeview(window,
                            columns=('file', 'filepath'),
                            show='headings')
output_table.heading('file',
                     text="Output File Name")
output_table.heading('filepath',
                     text="Output File Path")
output_table.column('file',
                    width=15)


# :::::::::::::::::: Labels :::::::::::::::::: 

# Label for Directory
directory_label_string = tk.StringVar()
output_labels_frame_label = tk.Label(output_files_label_frame,
                                     text="Files will be saved to:")
directory_label = tk.Label(output_files_label_frame,
                           textvariable=directory_label_string)

# Instructions Label
instructions_text1 = '1. Upload the files to be converted.\n\
2. Specify the output location.\n\
3. Specify output file type (txt/csv)'

instructions_text2 = 'Files will be generated in specified output location, \
labled as "(rolltype)_(filename).(outputtype)"\n \
Example: Uploaoding file: 5052 Q1.xls -> "regularhonorroll_5052Q1.txt"'


instructions_label1 = tk.Label(instructions_frame,
                               justify='left',
                               text=instructions_text1,
                               font=("Times New Roman", 14))
instructions_label2 = tk.Label(instructions_frame,
                               justify='center',
                               text=instructions_text2,
                               font=('Times New Roman', 12))


# Table Label

table_label = tk.Label(table_label_frame,
                       justify='left',
                       text='Selected Files',
                       font=('Times New Roman', 11))

# :::::::::::::::::: Events ::::::::::::::::::
# events
input_table.bind('<<TreeviewSelect>>', toggle_delete)
# table.bind('<Delete>', delete_items)


#  -------------- Packing --------------
# 1st Frame - Instructions
instructions_frame.pack()
instructions_label1.pack()
instructions_label2.pack(side='bottom')

# 2nd Frame - Upload button, directory button, output type radio buttons
setup_buttons_frame.pack()
get_files_button.pack(side='left')
set_directory_button.pack(side='left', padx=100)
radio_frame.pack(pady=20, side='left')
txt_radio_button.pack()
csv_radio_button.pack()

# 3rd Frame - Output location labels
output_files_label_frame.pack()
output_labels_frame_label.pack(side='left')
directory_label.pack(side='right')

# 4th Frame - Table Label
placeholder_frame.pack()
table_label_frame.pack(side='left')
table_label.pack(side='left')

# 5th Frame - Input Table
input_table_frame.pack(fill='both')
input_table.pack(padx=20, fill="x")

# 6th Frame
del_and_start_buttons_frame.pack()
start_button.pack(pady=10, side='left', expand=True)
file_delete_button.pack(expand=True, side='left')

# Output Table
output_table.pack(padx=20, fill='x')


# -------------- Run --------------
window.mainloop()
