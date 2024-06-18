# The purpose of this script is simple
# It is will grab the 4 files [Status report, Composition report, Remediation Report, and Sign-Off Report]
# It will use it the Testing SP excel file 
# It will add new tabs for each file that was added 

import pandas as pd
import tkinter as tk
from tkinter import filedialog

def Create_Testing_Report():
    Testing_File = filedialog.askopenfilename()
    empty_attestation_df = pd.DataFrame()
    Status_Report_File = filedialog.askopenfilename()
    Composition_Report_File = filedialog.askopenfilename()
    Remediation_Report_File = filedialog.askopenfilename()
    SignOff_Report_File = filedialog.askopenfilename()
    Post_Report_File = filedialog.askopenfilename()
    App_Name = input('Enter the App Name: ')

    try: 
       Testing_File = pd.read_excel(Testing_File)
       Status_Report_File = pd.read_csv(Status_Report_File)
       Composition_Report_File = pd.read_csv(Composition_Report_File)
       Remediation_Report_File = pd.read_csv(Remediation_Report_File)
       SignOff_Report_File = pd.read_csv(SignOff_Report_File)
       Post_Report_File = pd.read_csv(Post_Report_File)
       Application_name = App_Name

       Output_filename = f"{Application_name} Testing 1st half 2024.xlsx"
       with pd.ExcelWriter(Output_filename, engine='xlsxwriter') as writer:
           Testing_File.to_excel(writer, sheet_name="Testing", index=False)
           empty_attestation_df.to_excel(writer, sheet_name="Attestation", index=False)
           Status_Report_File.to_excel(writer, sheet_name="Status", index=False)
           Composition_Report_File.to_excel(writer, sheet_name="Composition", index=False)
           Remediation_Report_File.to_excel(writer, sheet_name="Remediation", index=False)
           SignOff_Report_File.to_excel(writer, sheet_name="Sign Off", index=False)
           Post_Report_File.to_excel(writer, sheet_name="POST", index=False)
       print(f"Processing is completed '{Application_name}' created successfully")

    except FileNotFoundError:
        print("File Not Found. Please Check the file paths.")
    except Exception as e:
        print(f"An error occurred: {e}")


Create_Testing_Report()
