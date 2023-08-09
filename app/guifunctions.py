from tkinter import filedialog
from tkinter import messagebox
import excel_converter as ec
import os
import errno


def start_button_state(start_button):
    if len(file_name) > 0 and output_locatin_selected and output_type_selected:
        start_button["state"] = tk.NORMAL
        start_button["fg"] = "green"
    else:
        start_button["state"] = tk.DISABLED
        start_button["bg"] = "white"


def get_files_func():
    global file_name
    file_name = list(
        filedialog.askopenfilenames(initialdir=get_files_path, title="Select File(s)")
    )
    max_file_path_column_length = 0
    for file in file_name:
        if file.endswith(".xls") or file.endswith(".xlsx"):
            max_file_path_column_length = max(max_file_path_column_length, len(file))
            insert_to_input_table(clean_string_for_table_insert(file), file)
        else:
            messagebox.showerror(
                "Warning!", "Only .xls and .xlsx files accepted!", parent=window
            )  # To do: figure out how to get this warning on the window itself

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
        file_delete_button["state"] = tk.NORMAL
        file_delete_button["fg"] = "red"
    else:
        file_delete_button["state"] = tk.DISABLED
        file_delete_button["bg"] = "white"
    start_button_state()


def delete_items(_):
    for i in input_table.selection():
        file_name.remove(input_table.item(i)["values"][1])
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
    return file_path_string.split("/")[-1].replace(" ", "").split(".")[0]


def make_new_dir(dir_path: str):
    os.makedirs(dir_path)


def get_output_files():
    return [
        fn
        for fn in os.listdir(output_location)
        if fn.lower().endswith(output_type)
        and os.path.isfile(os.path.join(output_location, fn))
    ]


def populate_output_table():
    output_files = get_output_files()
    print(output_files)
    for file in output_files:
        insert_to_output_table(
            clean_string_for_table_insert(file), f"{output_location}/{file}"
        )


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
