import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import numpy as np
import re
import xlwings as xw

# Declare df as a global variable
df = None

# initalise the tkinter GUI
root = tk.Tk()
root.title("eBay Transaction Report Consolidation")

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
            df.columns = df.iloc[10]
            df = df.iloc[1:]

            # Delete rows 1-11
            df = df.drop(df.index[0:10])

            # List of column names to be deleted
            columns_to_delete = ['Buyer username','Buyer name','Ship to city','Ship to province/region/state',
                                 'Ship to zip','Ship to country','Payout currency','Payout date','Payout ID',
                                 'Payout method','Payout status','Reason for hold','Item title','Custom label',
                                 'Transaction currency','Exchange rate']
            
            # Drop the specified columns
            df.drop(columns=columns_to_delete, inplace=True)

            # Convert the 'date_column' to datetime type
            df['Transaction creation date'] = pd.to_datetime(df['Transaction creation date'])

            # Format the datetime values to remove the time part
            df['Transaction creation date'] = df['Transaction creation date'].dt.strftime('%d-%m-%Y')
            
            # List of values to be removed
            values_to_remove = ['Payment dispute','Payout']

            # Filter rows with specified values in the 'Type' column
            rows_to_remove = df[df['Type'].isin(values_to_remove)]

            # Drop the filtered rows from the DataFrame
            df = df.drop(rows_to_remove.index)

            # Replace '--' with 'np.nan' in all columns
            df = df.applymap(lambda x: np.nan if x == '--' else x)

            # Convert values in the following columns to float
            columns_to_convert = ['Net amount','Item subtotal','Shipping and handling','Seller collected tax',
                                  'eBay collected tax','Final Value Fee - fixed','Final Value Fee - variable',
                                  'Very high "item not as described" fee','Below standard performance fee',
                                  'International fee','Gross transaction amount']
            
            # Convert specified columns to float
            for column in columns_to_convert:
                df[column] = pd.to_numeric(df[column], errors='coerce').fillna(0.0).astype(float)

            # Convert values in the following columns to str
            columns_to_convert = ['Item ID','Legacy order ID','Transaction ID','Type','Reference ID','Description']

            # Convert specified columns to string
            for column in columns_to_convert:
                df[column] = df[column].fillna('').astype(str)

            # Create a new DataFrame with rows where the 'Order number' column is equal to 'Empty' or 'NaN'
            df2 = df[(df['Order number'] == 'Empty') | (df['Order number'].isna())]
            
            # print data types
            print(df.dtypes)

            # Get the list of columns to strip
            cols = ['Legacy order ID','Item ID','Reference ID','Transaction ID','Description']

            # Strip whitespace from all columns
            df[cols] = df[cols].apply(lambda x: x.str.strip())
            
            # Filter the DataFrame to include only 'Order' type rows
            df_order = df[df['Type'] == 'Order']
            
            # Perform aggregation
            df_aggregated = df.groupby('Order number').agg({
                    'Transaction creation date': 'min',
                    'Type': ','.join,
                    'Legacy order ID': 'unique',
                    'Net amount': 'sum',
                    'Item ID': 'unique',
                    'Transaction ID': 'unique',
                    'Quantity': 'max',  
                    'Shipping and handling': 'sum',
                    'Seller collected tax': 'sum',
                    'eBay collected tax': 'sum',
                    'Final Value Fee - fixed': 'sum',
                    'Final Value Fee - variable': 'sum',
                    'Very high "item not as described" fee': 'sum',
                    'Below standard performance fee': 'sum',
                    'International fee': 'sum',
                    'Gross transaction amount': 'sum',
                    'Reference ID': ','.join,
                    'Description': ','.join
                    }).reset_index()
            
            # Calculate the subtotal for each 'Order number' based on the 'Type' condition
            subtotal_helper = df_order.groupby('Order number')['Item subtotal'].sum()
            
            # Update 'Item subtotal' in the aggregated DataFrame using the helper
            df_aggregated['Item subtotal'] = df_aggregated['Order number'].map(subtotal_helper)

            # Reset the index
            df_aggregated = df_aggregated.reset_index(drop=True)

            # Combine the two data frame
            df3 = pd.concat([df_aggregated, df2])
            
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

            df3['Item ID'] = df3['Item ID'].apply(clean_brackets_and_quotes)
            df3['Transaction ID'] = df3['Transaction ID'].apply(clean_brackets_and_quotes)
            df3['Legacy order ID'] = df3['Legacy order ID'].apply(clean_brackets_and_quotes)

            # Remove the .0 decimal in column Item ID and Transaction ID
            df3['Item ID'] = df3['Item ID'].str.replace('.0', '')
            df3['Transaction ID'] = df3['Transaction ID'].str.replace('.0', '')
            
            # Remove any leading or trailing space
            df3['Item ID'] = df3['Item ID'].str.strip(" ")
            df3['Transaction ID'] = df3['Transaction ID'].str.strip(" ")
            
            # Remove all spaces and special characters
            df3['Description'] = df3['Description'].str.replace(r"[^\d\-+\.\_]", "")
            df3['Reference ID'] = df3['Reference ID'].str.replace(r"[^\d\-+\.\_]", "")

            # Replace all multiple commas with a single comma
            df3['Description'] = df3['Description'].str.replace(r",,", ",")
            df3['Reference ID'] = df3['Reference ID'].str.replace(r",,", ",")

            # Remove any leading or trailing commas
            df3['Description'] = df3['Description'].str.strip(",")
            df3['Reference ID'] = df3['Reference ID'].str.strip(",")

            # Remove any leading or trailing spaces
            df3['Description'] = df3['Description'].str.strip(" ")
            df3['Reference ID'] = df3['Reference ID'].str.strip(" ")

            # Remove all spaces in the middle of the string
            df3['Reference ID'] = df3['Reference ID'].str.replace(r"\s+", "")
            df3['Description'] = df3['Description'].str.replace(r"\s+", "")
                
        
            new_order = ['Transaction creation date','Type','Order number','Legacy order ID','Net amount','Item ID',
               'Transaction ID','Quantity','Item subtotal','Shipping and handling','Seller collected tax',
               'eBay collected tax','Final Value Fee - fixed','Final Value Fee - variable','Very high "item not as described" fee',
               'Below standard performance fee','International fee','Gross transaction amount','Reference ID','Description']
            
            df3 = df3.reindex(columns=new_order)


            # Export the transformed data to a new Excel file
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", "eBay Transaction Report Consolidation.xlsx")
            df3.to_excel(file_path, index=False, freeze_panes=(1, 1))

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

