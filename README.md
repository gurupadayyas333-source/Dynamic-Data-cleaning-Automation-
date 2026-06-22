# Dynamic-Data-cleaning-Automation-
Dynamic Data Cleaning Automation Tool built with Python, Tkinter, and Pandas. Clean Excel files by removing duplicates, blank rows, trimming spaces, and formatting text with a user-friendly GUI.
# Dynamic Data Cleaning Automation GUI

## Overview

This project is a Python-based Graphical User Interface (GUI) application that automates common data-cleaning tasks for Excel files. Built using Tkinter and Pandas, the tool allows users to select an Excel file, choose desired cleaning operations, and save the cleaned output to a specified folder without writing any code.

## Features

* Browse and select any Excel file (.xlsx, .xls)
* Remove duplicate rows
* Remove blank rows
* Remove leading and trailing spaces from text columns
* Convert text columns to Title Case
* Select custom output folder
* Export cleaned data to a new Excel file
* Reset file paths and cleaning options
* Dynamic processing for any Excel dataset

## Technologies Used

* Python
* Tkinter (GUI Development)
* Pandas (Data Processing)
* OpenPyXL (Excel File Handling)

## Installation

1. Clone the repository:

```bash
git clone https://github.com/your-username/dynamic-data-cleaning-gui.git
```

2. Navigate to the project folder:

```bash
cd dynamic-data-cleaning-gui
```

3. Install dependencies:

```bash
pip install pandas openpyxl
```

## Usage

1. Run the Python script:

```bash
python data_cleaning_gui.py
```

2. Select an input Excel file.
3. Choose the desired data-cleaning operations.
4. Select an output folder.
5. Click **Clean Data**.
6. The cleaned Excel file will be automatically saved in the selected location.

## Learning Outcomes

This project demonstrates:

* GUI development with Tkinter
* File handling in Python
* Data cleaning using Pandas
* User-friendly automation solutions
* Dynamic Excel data processing

## Future Enhancements

* Remove special characters
* Standardize date formats
* Handle missing values with custom rules
* Column renaming functionality
* Data profiling dashboard

## Author

Developed as part of a Python Data Analytics learning project to automate repetitive data-cleaning tasks and improve productivity.
