import pandas as pd
from tkinter import filedialog

def Get_Total():
    Testing_File = filedialog.askopenfilename()
    Application_name = "Test#12.xlsx"

    try:
        Testing_File = pd.read_excel(Testing_File)
        Testing_File.loc['Total'] = f'=COUNTA(UNIQUE(A2:A{len(Testing_File)+1}))'
        Output_filename = f"{Application_name}"
        with pd.ExcelWriter(Output_filename, engine='xlsxwriter') as writer:
            Testing_File.to_excel(writer, sheet_name="Testing", index=False)

    except FileNotFoundError:
        print("File Not Found. Please Check the file paths.")

Get_Total()