import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# -------------------------------
# Functions
# -------------------------------

input_file = ""
output_folder = ""


def browse_input():
    global input_file
    input_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    input_label.config(text=input_file)


def browse_output():
    global output_folder
    output_folder = filedialog.askdirectory(
        title="Select Output Folder"
    )
    output_label.config(text=output_folder)


def clean_data():

    if not input_file:
        messagebox.showerror("Error", "Please select an input Excel file.")
        return

    if not output_folder:
        messagebox.showerror("Error", "Please select an output folder.")
        return

    try:
        # Read Excel
        df = pd.read_excel(input_file)

        # Remove Duplicate Rows
        if duplicate_var.get():
            df = df.drop_duplicates()

        # Remove Blank Rows
        if blank_var.get():
            df = df.dropna(how='all')

        # Remove Leading and Trailing Spaces
        if spaces_var.get():
            text_cols = df.select_dtypes(include='object').columns

            for col in text_cols:
                df[col] = df[col].astype(str).str.strip()

        # Convert Text Columns to Title Case
        if title_var.get():
            text_cols = df.select_dtypes(include='object').columns

            for col in text_cols:
                df[col] = df[col].astype(str).str.title()

        # Output File Name
        file_name = os.path.basename(input_file)
        output_file = os.path.join(
            output_folder,
            f"Cleaned_{file_name}"
        )

        # Save Excel
        df.to_excel(output_file, index=False)

        messagebox.showinfo(
            "Success",
            f"Cleaned file saved successfully!\n\n{output_file}"
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))


def reset_paths():
    global input_file, output_folder

    input_file = ""
    output_folder = ""

    input_label.config(text="No file selected")
    output_label.config(text="No folder selected")


def reset_checkboxes():
    duplicate_var.set(False)
    blank_var.set(False)
    spaces_var.set(False)
    title_var.set(False)


# -------------------------------
# GUI Window
# -------------------------------

root = tk.Tk()
root.title("Dynamic Data Cleaning Automation")
root.geometry("700x500")

# Input File
tk.Label(
    root,
    text="Input Excel File",
    font=("Arial", 10, "bold")
).pack(pady=5)

tk.Button(
    root,
    text="Browse Excel File",
    command=browse_input
).pack()

input_label = tk.Label(root, text="No file selected", wraplength=650)
input_label.pack(pady=5)

# Output Folder
tk.Label(
    root,
    text="Output Folder",
    font=("Arial", 10, "bold")
).pack(pady=5)

tk.Button(
    root,
    text="Browse Output Folder",
    command=browse_output
).pack()

output_label = tk.Label(root, text="No folder selected", wraplength=650)
output_label.pack(pady=5)

# Cleaning Options
tk.Label(
    root,
    text="Select Data Cleaning Options",
    font=("Arial", 12, "bold")
).pack(pady=10)

duplicate_var = tk.BooleanVar()
blank_var = tk.BooleanVar()
spaces_var = tk.BooleanVar()
title_var = tk.BooleanVar()

tk.Checkbutton(
    root,
    text="Remove Duplicate Rows",
    variable=duplicate_var
).pack(anchor="w", padx=50)

tk.Checkbutton(
    root,
    text="Remove Blank Rows",
    variable=blank_var
).pack(anchor="w", padx=50)

tk.Checkbutton(
    root,
    text="Remove Leading & Trailing Spaces (Text Columns)",
    variable=spaces_var
).pack(anchor="w", padx=50)

tk.Checkbutton(
    root,
    text="Convert Text Columns to Title Case",
    variable=title_var
).pack(anchor="w", padx=50)

# Buttons Frame
button_frame = tk.Frame(root)
button_frame.pack(pady=20)

tk.Button(
    button_frame,
    text="Clean Data",
    width=15,
    bg="lightgreen",
    command=clean_data
).grid(row=0, column=0, padx=10)

tk.Button(
    button_frame,
    text="Reset Paths",
    width=15,
    command=reset_paths
).grid(row=0, column=1, padx=10)

tk.Button(
    button_frame,
    text="Reset Checkboxes",
    width=15,
    command=reset_checkboxes
).grid(row=0, column=2, padx=10)

# Run GUI
root.mainloop()