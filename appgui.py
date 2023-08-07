import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

# -------------- Functions --------------

file_name = ""
def get_files_func():
    global file_name
    file_name = filedialog.askopenfilenames(initialdir="C:/Users/ariel/Desktop/Python Projects/Excel/HonourRoll2",title="Select A File")
    for file in file_name:
        insert_to_table(clean_string_for_table_insert(file))


def clean_string_for_table_insert(string_to_clean):
    return string_to_clean.split("/")[-1].replace(" ","")


def insert_to_table(string_to_insert):
    table.insert(parent="",index=tk.END,values=(string_to_insert))


def toggle_delete(_):
    print("Event Triggered")
    # if len(file_name) > 0:
    #     file_delete_button['state'] = "normal"
    # file_delete_button['state'] = "disabled"
    # pass

# -------------- Window --------------

window = tk.Tk()
window.title("Val's App")
window.geometry("800x500")


# -------------- Widgets --------------

# Get File Button
get_files_button = ttk.Button(window,text="Get Files",command=get_files_func)



# File Window
table = ttk.Treeview(window,columns=('FileName'),show='headings')
table.heading('FileName',text="File Name")
table.selection_remove(*table.selection())
table.bind('<<TreeviewSelect>>', toggle_delete)


# File Delete Button
file_delete_button = ttk.Button(window,text="Delete",state="disabled")






# -------------- Packing --------------
get_files_button.pack(pady=20)
table.pack(padx=15,fill="x")
file_delete_button.pack()

# -------------- Run --------------
window.mainloop()


# Testing shit
print(file_name)
