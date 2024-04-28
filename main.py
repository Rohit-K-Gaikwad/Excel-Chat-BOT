import pandas as pd
import webbrowser
from tkinter import *


URL_PATH = "your_excel_sheet_path" # Paste your excel sheet path here.

# Read Excel data into DataFrame
df = pd.read_excel(URL_PATH)


# Function to open link in browser
def open_link():
    selected_student = student_var.get()
    marksheet_link = df.loc[
        df["File Name"] == selected_student, "File Access Link"
    ].iloc[0]
    webbrowser.open_new_tab(marksheet_link)


# Initialize Tkinter
root = Tk()
root.title("MarkSheet Bot")

# Student Dropdown
Label(root, text="Select Student:").grid(row=0, column=0)
student_var = StringVar(root)
students_dropdown = OptionMenu(root, student_var, *df["File Name"])
students_dropdown.grid(row=0, column=1)

# Button to open marksheet
open_button = Button(root, text="Open Marksheet", command=open_link)
open_button.grid(row=1, column=0, columnspan=2)

# Start Tkinter event loop
root.mainloop()
