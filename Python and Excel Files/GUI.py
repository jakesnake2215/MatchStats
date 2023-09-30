import tkinter as tk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import openpyxl
import os
import shutil

def dragAndDrop():
    processed_file_path = []  # Create a mutable object to store the processed file path

    def handle_drop(event):
        file_path = event.data
        lb.insert(tk.END, file_path)
        processed_file_path.append(process_file(file_path))  # Call a function to process the file path
        root.destroy()  # Close the window when the file is dropped

    root = TkinterDnD.Tk()

    lb = tk.Listbox(root)
    lb.insert(1, "drag files to here")

    # register the listbox as a drop target
    lb.drop_target_register(DND_FILES)
    lb.dnd_bind('<<Drop>>', handle_drop)

    lb.pack()

    def process_file(file_path):
        # Add your file processing logic here
        # Output the file path as a string
        return file_path  # Example processing, modify as needed

    root.mainloop()
    return processed_file_path[0]  # Return the first (and only) item in the list


def openExcelFile():
    file_path = "C:\\Users\\jakeg\\OneDrive\\Desktop\\r6-dissect-v0.11.1-windows-amd64\\Stats.xlsx"  # Specify the pre-pathed Excel file
    if os.path.exists(file_path):
        os.system("start excel {}".format(file_path))  # Open the Excel file
    else:
        print("File not found:", file_path)

    
# root = tk.Tk()
# root.title("Open Prepathed Excel File")

# button = tk.Button(root, text="Open Excel File", command=openExcelFile)
# button.pack(padx=20, pady=20)

# root.mainloop()

processed_file_path = dragAndDrop()


# Find the index of the last backslash
last_backslash_index = processed_file_path.rfind("/")

# # Extract text after the last backslash
text_after_last_backslash = processed_file_path[last_backslash_index + 1:]
destination_path = "C:\\Users\\jakeg\\OneDrive\\Pictures\\Clips\\" + text_after_last_backslash
shutil.copytree(processed_file_path, destination_path)
# print("Processed File Path:", processed_file_path)
