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
root.title("Amazon Transaction Report Consolidation")

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
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

            # Delete the header row
            df.columns = df.iloc[6]
            df = df.iloc[1:]

            # Delete rows 1-7
            df = df.drop(df.index[0:6])
            
            # List of values to be removed
            values_to_remove = ['Transfer']

            # Filter rows with specified values in the 'Type' column
            rows_to_remove = df[df['type'].isin(values_to_remove)]

            # Drop the filtered rows from the DataFrame
            df = df.drop(rows_to_remove.index)

            df['marketplace'] = df['marketplace'].str.replace('Amazon.com', 'amazon.com')

            # Convert the 'date/time' column to the datetime format
            df['date/time'] = pd.to_datetime(df['date/time'])

            # Create a new DataFrame with rows where the 'Order number' column is equal to 'Empty' or 'NaN'
            df1 = df[(df['order id'] == 'Empty') | (df['order id'].isna())]

            # Perform aggregation
            df_aggregated = df.groupby('order id').agg({
                            'date/time': 'min',
                            'settlement id': 'unique',
                            'type': ','.join,
                            'sku': 'unique',
                            'description': 'unique',
                            'quantity': 'unique',
                            'marketplace': 'unique',
                            'account type': 'unique',
                            'fulfillment': 'unique',
                            'order city': 'unique',
                            'order state': 'unique',
                            'order postal': 'unique',
                            'tax collection model': 'unique',
                            'product sales': 'sum',
                            'product sales tax': 'sum',
                            'shipping credits': 'sum',
                            'shipping credits tax': 'sum',
                            'gift wrap credits': 'sum',
                            'giftwrap credits tax': 'sum',
                            'Regulatory Fee': 'sum',
                            'Tax On Regulatory Fee': 'sum',
                            'promotional rebates': 'sum',
                            'promotional rebates tax': 'sum',
                            'marketplace withheld tax': 'sum',
                            'selling fees': 'sum',
                            'fba fees': 'sum',
                            'other transaction fees': 'sum',
                            'other': 'sum',
                            'total': 'sum'
                            }).reset_index()
            
            # Reset the index
            df_aggregated = df_aggregated.reset_index(drop=True)

            # Combine the two data frame
            df3 = pd.concat([df_aggregated, df1])

            new_order = ['date/time','settlement id','type','order id','sku','description','quantity',
                         'marketplace','account type','fulfillment','order city','order state','order postal',
                         'tax collection model','product sales','product sales tax','shipping credits',
                         'shipping credits tax','gift wrap credits','giftwrap credits tax','Tax On Regulatory Fee',
                         'promotional rebates','promotional rebates tax','marketplace withheld tax',
                         'selling fees','fba fees','other transaction fees','other','total']
            
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
            
            df3['date/time'] = df3['date/time'].apply(clean_brackets_and_quotes)
            df3['settlement id'] = df3['settlement id'].apply(clean_brackets_and_quotes)
            df3['type'] = df3['type'].apply(clean_brackets_and_quotes)   
            df3['order id'] = df3['order id'].apply(clean_brackets_and_quotes)
            df3['sku'] = df3['sku'].apply(clean_brackets_and_quotes)  
            df3['description'] = df3['description'].apply(clean_brackets_and_quotes)   
            df3['quantity'] = df3['quantity'].apply(clean_brackets_and_quotes)  
            df3['marketplace'] = df3['marketplace'].apply(clean_brackets_and_quotes) 
            df3['account type'] = df3['account type'].apply(clean_brackets_and_quotes) 
            df3['fulfillment'] = df3['fulfillment'].apply(clean_brackets_and_quotes) 
            df3['order city'] = df3['order city'].apply(clean_brackets_and_quotes)  
            df3['order state'] = df3['order state'].apply(clean_brackets_and_quotes)  
            df3['order postal'] = df3['order postal'].apply(clean_brackets_and_quotes) 
            df3['tax collection model'] = df3['tax collection model'].apply(clean_brackets_and_quotes)
            
            # Remove 'nan' in the following columns
            df3['date/time'] = df3['date/time'].str.replace('nan', '')
            df3['settlement id'] = df3['settlement id'].str.replace('nan', '')
            df3['type'] = df3['type'].str.replace('nan', '')
            df3['order id'] = df3['order id'].str.replace('nan', '')
            df3['sku'] = df3['sku'].str.replace('nan', '')
            df3['description'] = df3['description'].str.replace('nan', '')
            df3['quantity'] = df3['quantity'].str.replace('nan', '')
            df3['marketplace'] = df3['marketplace'].str.replace('nan', '')
            df3['account type'] = df3['account type'].str.replace('nan', '')
            df3['fulfillment'] = df3['fulfillment'].str.replace('nan', '')
            df3['order city'] = df3['order city'].str.replace('nan', '')
            df3['order state'] = df3['order state'].str.replace('nan', '')
            df3['order postal'] = df3['order postal'].str.replace('nan', '')
            df3['tax collection model'] = df3['tax collection model'].str.replace('nan', '')
 
            # Remove 'Amazon.com'
            df3['marketplace'] = df3['marketplace'].str.replace('Amazon.com', '')
            
            # Convert the 'date/time' column into the datetime format
            df3['date/time'] = pd.to_datetime(df3['date/time'])

            # Format the datetime objects in the 'date/time' column into the desired format
            df3['date/time'] = df3['date/time'].dt.strftime('%B %d, %Y %I:%M:%S %p PDT')  

            # Remove the double quotation marks in the 'marketplace' column
            df3['description'] = df3['description'].str.replace('"', '')       

            # Export the transformed data to a new Excel file
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", "Amazon Transaction Report Consolidation.xlsx")
            df3.to_excel(file_path, index=False, freeze_panes=(1, 4))


    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    

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