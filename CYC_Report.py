import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import numpy as np
import re
import xlwings as xw
import customtkinter
from datetime import datetime

# Declare df as a global variable
df = None

# Setting up theme of the app
customtkinter.set_appearance_mode("system")

# Setting up them of your components
customtkinter.set_default_color_theme("blue")

# initalise the tkinter GUI
root = customtkinter.CTk()
root.title("CYC Report")

root.geometry("450x450") # set the root dimensions
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.

# Frame for TreeView
frame1 = tk.LabelFrame(root, text="Excel Data",  bg="lightgrey", fg="black", font=("Arial", 8, "bold"))
frame1.place(height=200, width=550, rely=0, relx=0.01)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File", bg="lightgrey", fg="black", font=("Arial", 8, "bold"))
file_frame.place(height=170, width=480, rely=0.65, relx=0)

# Buttons
button1 = customtkinter.CTkButton(file_frame, text="Browse for File", command=lambda: File_dialog(), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button1.place(rely=0.2, relx=0.01)

button2 = customtkinter.CTkButton(file_frame, text="Run Transformation", command=lambda: Transform_1(), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button2.place(rely=0.6, relx=0.01)

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected", background="lightgrey", foreground="blue", font=("Arial", 8, "bold"))
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
    
    if not os.path.isfile(file_path):
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename, engine=None)
        else:
            df = pd.read_excel(excel_filename, engine=None)
    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    
    # Delete the header row
    df.columns = df.iloc[3]
    df = df.iloc[2:]
   
    # Delete rows 1-4
    df = df.drop(df.index[0:3])
    
    # Drop rows where any cell has value 'Jasper'
    df = df[~df.isin(['Jasper']).any(axis=1)]
    
    # Remove leading and trailing spaces
    df = df[(df['Status'].str.strip() != 'Grand Total:') & (df['Status'].str.strip() != 'Sub Total:')]

    # Drop rows where the column has value 'Grand Total' or 'Subtotal'
    df = df[(df['Status'] != 'Grand Total:') & (df['Status'] != 'Sub Total:')]

    # Add a new column 'Part Number' to df
    df['Part Number'] = None

    # Take the Part Number data and fill in the column Part Number
    mask = df['SO #'].str.contains(' ', na=False)
    df.loc[mask, ['Part Number', 'SO #']] = df.loc[mask, ['SO #', 'Part Number']].values

    df['SO #'] = df['SO #'].fillna(method='ffill')
    df['Customer PO'] = df['Customer PO'].fillna(method='ffill')

    # Convert the 'Date Issued' column to string and split it
    df['Date Issued'] = df['Date Issued'].apply(lambda x: str(x).split(' ')[0])

    # Convert the 'Date Fulfilled' column to string and split it
    df['Date Fulfilled'] = df['Date Fulfilled'].apply(lambda x: str(x).split(' ')[0])

    # Replace variations of 'nan' with 'empty' in all columns
    df = df.applymap(lambda x: '' if str(x).lower() == 'nan' else x)

    # List of columns you want to modify
    columns_to_modify = ['Ordered', 'Shipped', 'Remaining']

    # # Replace negative values with their absolute values in the specified columns
    # for column in columns_to_modify:
    #     df[column] = df[column].apply(lambda x: abs(x))

    # Format the valida values date from column Date Fulfilled and Date Issued
    def convert_date(date_str):
       try:
           return datetime.strptime(date_str, '%Y-%m-%d').strftime('%d/%m/%Y')
       except ValueError:
           return date_str  # return original value

    df['Date Fulfilled'] = df['Date Fulfilled'].apply(convert_date)
    df['Date Issued'] = df['Date Issued'].apply(convert_date)
    
    # Arrange the columns base on the new_order list
    new_order = ['SO #', 'Part Number', 'Customer', 'Customer PO', 'Date Issued', np.nan, 'Date Fulfilled',
                  'Status', 'Ordered',  'Shipped', 'Remaining' ]
    
    df = df.reindex(columns=new_order)

    print(df.dtypes)

    # Export the transformed data to a new Excel file
    file_path = os.path.join(os.path.expanduser("~"), "Desktop", "CYC_Report.xlsx")
    df.to_excel(file_path, index=False, freeze_panes=(1, 0))
    messagebox.showinfo("Success", "File has been successfully written.")

    # Open the Excel file and set all columns width to 15
    with xw.App(visible=False) as app:
        wb = xw.Book(file_path)

        # Loop through all worksheets in the workbook
        for ws in wb.sheets:
            # Loop through all columns in the worksheet
            for column in ws.api.UsedRange.Columns:
                column.ColumnWidth = 15

        # Save the workbook if needed
        wb.save()

        # Close the workbook
        wb.close()
        
    
    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # let the column heading = column name

    df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
    return None
    
def clear_data():
    tv1.delete(*tv1.get_children())
    return None

root.mainloop()

