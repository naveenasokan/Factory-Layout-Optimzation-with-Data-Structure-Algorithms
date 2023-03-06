from dataclasses import replace
from posixpath import split
from telnetlib import theNULL
import tkinter
from tkinter import filedialog
from tkinter.filedialog import asksaveasfilename
import pandas as pd
import os,sys

root = tkinter.Tk()
filez = filedialog.askopenfilenames(parent=root, title='Choose a file',filetypes=[("Excel files", "*.xlsx")])
basedirectory=os.path.basename(os.path.dirname(filez[0]))
if not filez:
    sys.exit()
name=[os.path.basename(f1) for f1 in filez]
name=[f1.replace(f"{basedirectory}_data_","") for f1 in name]
name=[f1.replace(".xlsx","") for f1 in name]
name=f"{basedirectory}_data_" + "_".join(name)



testdf=pd.read_excel(filez[0],None)
sheetnames=testdf.keys()
df_data_merged={}

for file in filez:

    data=pd.read_excel(file,None)
    for k,sheet in enumerate(sheetnames):
        if sheet in df_data_merged:
            df_data_merged[sheet]=df_data_merged[sheet].append(data[sheet])
        else:
             df_data_merged[sheet]= data[sheet]




    
# Close the Pandas Excel writer and output the Excel file.
f=asksaveasfilename(initialfile = f"{name}.xlsx",
defaultextension=".xlsx",filetypes=[("Excel files", "*.xlsx")])
if not f:
    sys.exit()
writer = pd.ExcelWriter(f)

for keys in df_data_merged:

# Write each dataframe to a different worksheet.
    df_data_merged[keys].to_excel(writer, sheet_name=keys,index=False)
writer.save()