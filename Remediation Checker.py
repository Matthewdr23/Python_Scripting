'''
The purpose of this code will be to add 2 files the revokes and the POST report file and it will automate the column of vlookup and the concat columns so that the only this will be to manually double check the first 
couple of times until this script will be the go to script to check 

1. The script will ask for the remediation report 
2. The script will ask for the POST Status report 
3. The script will indicate that the script is running 
4. It will produce a file called <Application Name> Prove Out
5. Check the Prove out and make sure everything is correct
6. Set the status of the recert as ready for 1st line

'''
import pandas as pd
import tkinter as tk
from tkinter import filedialog


remediation_filename = ""
Post_filename = ""

    # Process the files and create something
    # (similar to the PySimpleGUI example)

def update_file_paths():
    Remediation_Report_Path = Remediation_Report_Var.get()
    Post_Report_Path = Post_Report_Var.get()
    Remediation_Report_Text.delete(1.0, tk.END)
    Post_Report_Text.delete(1.0, tk.END)
    Remediation_Report_Text.insert(tk.END, Remediation_Report_Path)
    Post_Report_Text.insert(tk.END, Post_Report_Path)

def empty_file_path():
    Remediation_Report_Var.set("")
    Post_Report_Var.set("")
    Remediation_Report_Text.delete(1.0, tk.END)
    Post_Report_Text.delete(1.0, tk.END)


root = tk.Tk()
root.title("Prove Out Generator")

# Name entry
name_label = tk.Label(root, text="Enter Application Name:")
name_entry = tk.Entry(root)
App_name = name_entry.get()
name_label.grid(row=0, column=0, padx=10, pady=10)
name_entry.grid(row=0, column=1, padx=10, pady=10)

# File selection buttons
Remediation_Report_Var = tk.StringVar()
Post_Report_Var = tk.StringVar()

def select_remediation_report():
    #global remediation_filename
    remediation_filename = filedialog.askopenfilename()
    Remediation_Report_Var.set(remediation_filename)
    update_file_paths()  # Update the text box

def select_post_report():
    #global Post_filename
    Post_filename = filedialog.askopenfilename()
    Post_Report_Var.set(Post_filename)
    update_file_paths()  # Update the text box

App_name = name_entry.get()
print(App_name)
Remediation_Report_Button = tk.Button(root, text="Select Remediation Report", command=select_remediation_report)
Post_Report_Button = tk.Button(root, text="Select POST Report", command=select_post_report)
global remediation_Report_Location
global post_Report_Location
remediation_Report_Location = remediation_filename
Post_Report_Location = Post_filename
print(remediation_Report_Location)
print(Post_Report_Location)
def Create_Report():

    try:
        remediation_Report = pd.read_excel(remediation_Report_Location)
        Post_report = pd.read_excel(Post_Report_Location)

        # Get the application name from the user
        Application_Name = App_name

        FoundInPOST_Function = "=VLOOKUP(C2,POST!C:D,2,FALSE)"

        # Add the new column to the 'POST' dataframe
        remediation_Report["Found in POST"] = FoundInPOST_Function

        # Create an Excel writer using XlsxWriter as the engine
        output_filename = f"{Application_Name} Prove Out.xlsx"
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            # Write data from file1 to a separate worksheet
            remediation_Report.to_excel(writer, sheet_name='Remediation', index=False)

            # Write data from file2 to another worksheet
            Post_report.to_excel(writer, sheet_name='POST', index=False)

        print(f"Excel file '{output_filename}' created successfully!")
        empty_file_path()

    except FileNotFoundError:
        print("File not found. Please check the file paths.")
    except Exception as e:
        print(f"An error occurred: {e}")

Remediation_Report_Button.grid(row=1, column=0, padx=10, pady=10)
Post_Report_Button.grid(row=2, column=0, padx=10, pady=10)

# Read-only text boxes for file paths
Remediation_Report_Text = tk.Text(root, height=1, width=80)
Post_Report_Text = tk.Text(root, height=1, width=80)

Remediation_Report_Text.grid(row=1, column=1, padx=10, pady=10)
Post_Report_Text.grid(row=2, column=1, padx=10, pady=10)

# Create button
create_button = tk.Button(root, text="Create", command=Create_Report)
create_button.grid(row=3, column=1, padx=10, pady=10, sticky="e")  # Move the button to the right

# Update file paths when selection changes
Remediation_Report_Var.trace_add("write", lambda *args: update_file_paths())
Post_Report_Var.trace_add("write", lambda *args: update_file_paths())



# Read data from the two input Excel files


root.mainloop()






