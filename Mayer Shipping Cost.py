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
root.title("Mayer Shipping Cost")

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
            df = pd.read_excel(excel_filename, engine=None, chunksize=10000)
    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    
    # Assuming df is your DataFrame and 'column' is the column you want to modify
    df['Meyer Part'] = df['Meyer Part'].apply(lambda x: x[:3] + '-' + x[3:])

    df = df.rename(columns={'Meyer Part': 'Inventory Number'})


    # Assuming df is your DataFrame
    df['Your Price'] = pd.to_numeric(df['Your Price'], errors='coerce')


    def calculate_mayer_shipping_cost(row):
      if pd.isnull(row['Your Price']):
         return ""
      else:
        description = str(row['Description']).strip()
        ltl = str(row['LTL']).lower().strip() == 'false'
        oversize = str(row['Oversize']).lower().strip() == 'no'
        oversize2 = str(row['Oversize']).lower().strip() == 'yes'
        addtl_handling_charge = str(row['Addtl Handling Charge']).lower().strip() == 'no'
        addtl_handling_charge2 = str(row['Addtl Handling Charge']).lower().strip() == 'yes'
        kit = '(kit)' in description.lower()
    
        

        if row['Your Price'] < 250:
            base = 10 if ltl else 5 + 200 
            oversize = 0 if oversize else 70
            addtl = 0 if addtl_handling_charge else 5
            kit = 10 if kit else 0
            if oversize2 and addtl_handling_charge2 and ltl:
                return base + 70

        elif row['Your Price'] >= 500:
            base = 25 if ltl else 20 + 200 
            oversize = 0 if oversize else 70
            addtl = 0 if addtl_handling_charge else 5
            kit = 25 if kit else 0
            if oversize2 and addtl_handling_charge2 and ltl:
                return base + 70

        else:
            base = 15 if ltl else 10 + 200
            oversize = 0 if oversize else 70
            addtl = 0 if addtl_handling_charge else 5 
            kit = 15 if kit else 0
            if oversize2 and addtl_handling_charge2 and ltl:
                return base + 70


        return base + oversize + addtl + kit

    df['Mayer Shipping Cost'] = df.apply(calculate_mayer_shipping_cost, axis=1)
  
    print(df.dtypes)

    # Export the transformed data to a new Excel file
    file_path = os.path.join(os.path.expanduser("~"), "Desktop", "Mayer Shipping Cost.xlsx")
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

