import pandas as pd
import tkinter as tk
from tkinter import filedialog

#Z:\\File Share ACLS\2024\File Shares 1st2024\1st2024 File Shares and Program Locations Package\Business Areas ACLS\CorpApps FS\Planning Analytics FS\Reports\Post\ERPADTP3-RSVDATA 1.xlsx

def Create_Report():

    # Get files
    #Step 1: Get the remediaton Report 
    #remediation_Report_Location = input("Enter Remediation File Location: ")
    remediation_Report_Location = filedialog.askopenfilename(title="Select Remediation Report")
    #Step 2: Get the POST Report
    post_Report_Location = filedialog.askopenfilename(title="Select POST Report")
    #Step 3: Enter the App Name
    # The only things that is needed it the name because the Prove out part of the name is automated
    App_name = input("Enter the Prove Out Name:") # Example Planning analytics, ResQ, etc


    try:
        print(remediation_Report_Location)
        remediation_Report = pd.read_csv(remediation_Report_Location)
        Post_report = pd.read_csv(post_Report_Location)

        # Get the application name from the user
        Application_Name = App_name
        #Make sure once you open the file that you fix the column so that not all the cells are lookinig at cell C2 but all the rows
        FoundInPOST_Function = "=VLOOKUP(C2,POST!C:D,2,FALSE)"
        #CONCAT_Method = "=concat()"

        # Add the new column to the 'POST' dataframe
        remediation_Report["Found in POST"] = FoundInPOST_Function
        #add a new column to concat 
        #remediation_Report["I+K"] = CONCAT_Method


        # Create an Excel writer using XlsxWriter as the engine
        output_filename = f"{Application_Name} Prove Out.xlsx"
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            # Write data from file1 to a separate worksheet
            remediation_Report.to_excel(writer, sheet_name='Remediation', index=False)

            # Write data from file2 to another worksheet
            Post_report.to_excel(writer, sheet_name='POST', index=False)

        print(f"Excel file '{output_filename}' created successfully!")
        

    except FileNotFoundError:
        print("File not found. Please check the file paths.")
    except Exception as e:
        print(f"An error occurred: {e}")


Create_Report()