import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import numpy as np
import re

# Declare df as a global variable
df = None

# initalise the tkinter GUI
root = tk.Tk()
root.title("Excel Automation Application")

root.geometry("600x600") # set the root dimensions
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.

# Frame for TreeView
frame1 = tk.LabelFrame(root, text="Excel Data")
frame1.place(height=300, width=600)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File")
file_frame.place(height=200, width=600, rely=0.65, relx=0)

# Buttons
button1 = tk.Button(file_frame, text="Browse for File", command=lambda: File_dialog())
button1.place(rely=0.2, relx=0.01)

button2 = tk.Button(file_frame, text="Run Transformation", command=lambda: Transform_1())
button2.place(rely=0.4, relx=0.01)

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)

# Treeview Widget
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).

treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview) # command means update the yaxis view of the widget
treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview) # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign the scrollbars to the Treeview Widget
treescrollx.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
treescrolly.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget

def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file["text"] = filename
    return None


def Transform_1():
    """If the file selected is valid this will load the file into the Treeview"""
    global df  # Declare df as global to update it
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename, engine=None)
        else:
            df = pd.read_excel(excel_filename, engine=None)
        
        # Removing the "-1" from column [Original Ref#3/PO Number]
        def g(df):
            df["Original Ref#3/PO Number"] = df["Original Ref#3/PO Number"].str.replace("-1", "")
            return df
        df = g(df.copy())
        
        # Convert column values into float
        df["Net Charge Amount"] = pd.to_numeric(df["Net Charge Amount"], errors='coerce').astype(float)

        # print data types
        print(df.dtypes)

        # [Original Ref#3/PO Number] blanks pull out (put back at end of process)
        df1 = df[(df['Original Ref#3/PO Number'] == 'Empty') | (df['Original Ref#3/PO Number'].isna())]

        # Perform aggregation
        df_aggregated = df.groupby('Original Ref#3/PO Number').agg({
                        'Invoice Date': 'unique',
                        'Original Customer Reference': 'unique',
                        'Original Ref#2': "unique",
                        'Original Department Reference Description': 'unique',
                        'RMA#': 'unique',
                        'Net Charge Amount': 'sum',
                        'Shipper Company': 'unique',
                        'Express or Ground Tracking ID': 'unique',
                        'Recipient Name': 'unique',
                        'Recipient Company': 'unique',
                        'Shipper Address Line 1': 'unique',
                        'Shipper Address Line 2': 'unique',
                        'Shipper City': 'unique',
                        'Shipper State': 'unique',
                        'Shipper Zip Code': 'unique'
                        }).reset_index()
        

        # Combine the two data frame
        df3 = pd.concat([df_aggregated, df1])

        new_order = ['Invoice Date','Original Ref#3/PO Number','Original Customer Reference','Original Ref#2',
                     'Original Department Reference Description','RMA#','Net Charge Amount','Shipper Company',
                     'Express or Ground Tracking ID','Recipient Name','Recipient Company','Shipper Address Line 1',
                     'Shipper Address Line 2','Shipper City','Shipper State','Shipper Zip Code']
        
        df3 = df3.reindex(columns=new_order)
        
        
        # Removing unwated strings 
        def clean_brackets_and_quotes(data):

                # Convert the data to a string.
                data = str(data)

                # Escape single quotes and double quotes in the pattern.
                pattern = r"^\[|\]|\'|\\"""
                
                # Remove square brackets and single quotes at the start and end of the string.
                data = re.sub(pattern, "", data)

                # Remove decimal point ".0" from the end of the string.
                data = re.sub(r"\.0$", "", data)

                # Remove all decimal points from the string.
                data = re.sub(r"\.$", "", data)

                return data
                
        df3['Invoice Date'] = df3['Invoice Date'].apply(clean_brackets_and_quotes)
        df3['Original Ref#3/PO Number'] = df3['Original Ref#3/PO Number'].apply(clean_brackets_and_quotes)
        df3['Original Customer Reference'] = df3['Original Customer Reference'].apply(clean_brackets_and_quotes)
        df3['Original Ref#2'] = df3['Original Ref#2'].apply(clean_brackets_and_quotes)
        df3['Original Department Reference Description'] = df3['Original Department Reference Description'].apply(clean_brackets_and_quotes)
        df3['RMA#'] = df3['RMA#'].apply(clean_brackets_and_quotes)
        df3['Net Charge Amount'] = df3['Net Charge Amount'].apply(clean_brackets_and_quotes)
        df3['Shipper Company'] = df3['Shipper Company'].apply(clean_brackets_and_quotes)
        df3['Express or Ground Tracking ID'] = df3['Express or Ground Tracking ID'].apply(clean_brackets_and_quotes)
        df3['Recipient Name'] = df3['Recipient Name'].apply(clean_brackets_and_quotes)
        df3['Recipient Company'] = df3['Recipient Company'].apply(clean_brackets_and_quotes)
        df3['Shipper Address Line 1'] = df3['Shipper Address Line 1'].apply(clean_brackets_and_quotes)
        df3['Shipper Address Line 2'] = df3['Shipper Address Line 2'].apply(clean_brackets_and_quotes)
        df3['Shipper City'] = df3['Shipper City'].apply(clean_brackets_and_quotes)
        df3['Shipper State'] = df3['Shipper State'].apply(clean_brackets_and_quotes)
        df3['Shipper Zip Code'] = df3['Shipper Zip Code'].apply(clean_brackets_and_quotes)
        
        # Replace variations of 'nan' with 'empty' in all columns
        df3 = df3.applymap(lambda x: '' if str(x).lower() == 'nan' else x)

        # Replace double quotes in all columns
        df3 = df3.applymap(lambda x: x.replace('"', '') if isinstance(x, str) else x)

        # Convert the values into float under column ['Net Charge Amount']
        df3["Net Charge Amount"] = pd.to_numeric(df3["Net Charge Amount"], errors='coerce').astype(float)

        # Convert the values of column ['Net Charge Amount'] into 2 decimal places
        def round_to_one_decimal(x):
            return round(x, 2)

        df3['Net Charge Amount'] = df3['Net Charge Amount'].apply(round_to_one_decimal)

        # Export the transformed data to a new Excel file
        file_path = os.path.join(os.path.expanduser("~"), "Desktop", "data1.xlsx")
        df1.to_excel(file_path, index=False, freeze_panes=(1, 1))

        # Export the transformed data to a new Excel file
        file_path = os.path.join(os.path.expanduser("~"), "Desktop", "FedEx Charges.xlsx")
        df3.to_excel(file_path, index=False, freeze_panes=(1, 2))

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None

root.mainloop()