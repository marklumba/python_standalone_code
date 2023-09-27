
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
from tkinter import messagebox
import os
import xlwings as xw
from datetime import datetime



# Declare df1 and df2 as global variables
df1 = None
df2 = None

# initalise the tkinter GUI
root = tk.Tk()
root.title("Filter and Matching Tool")

root.geometry("650x650") # set the root dimensions
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.


frame1 = tk.LabelFrame(root, text="Functions")
frame1.place(height=150, width=635, relx=0.01)

file_frame = tk.LabelFrame(root, text="Display File")
file_frame.place(height=200, width=635, rely=0.65, relx=0.01)

file_frame1 = tk.LabelFrame(root, text="Display File")
file_frame1.place(height=200, width=635, rely=0.30, relx=0.01)

# Buttons
button1 = tk.Button(file_frame, text="Select Custom Fitment File", command=lambda: File_dialog_1())
button1.place(rely=0.2, relx=0.01)

button2 = tk.Button(file_frame, text="Read The File into Data Frame", command=lambda: Read_data_1())
button2.place(rely=0.4, relx=0.01)

button3 = tk.Button(file_frame1, text="Select eBay Pre-Filter MVL File", command=lambda: File_dialog_2())
button3.place(rely=0.2, relx=0.01)

button4 = tk.Button(file_frame1, text="Read The File into Data Frame", command=lambda: Read_data_2())
button4.place(rely=0.4, relx=0.01)

button5 = tk.Button(frame1, text="Run Matching Filter", command=lambda: Run_Matching_Filter(df1, df2))
button5.place(rely=0.2, relx=0.01)

# button6 = tk.Button(frame1, text="Check Compatibility", command=lambda: Check_Compatibility(df1, df2))
# button6.place(rely=0.45, relx=0.01)

# Check Compatibility button
button6 = tk.Button(frame1, text="Check Compatibility", command=lambda: Check_Compatibility())
button6.place(rely=0.45, relx=0.01)



label_file1 = ttk.Label(file_frame, text="No File Selected")
label_file1.place(rely=0, relx=0)

label_file2 = ttk.Label(file_frame1, text="No File Selected")
label_file2.place(rely=0, relx=0)

def File_dialog_1():
    """This Function will open the file explorer and assign the chosen file path to label_file1"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file1["text"] = filename
    return None

def File_dialog_2():
    """This Function will open the file explorer and assign the chosen file path to label_file2"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file2["text"] = filename
    return None

def Read_data_1():
    """If the file selected is valid this will load the file into the Treeview"""
    global df1  # Declare df as global to update it
    file_path = label_file1["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df1 = pd.read_csv(excel_filename)
        else:
            df1 = pd.read_excel(excel_filename)

            # Convert numeric columns in df1 to object
            numeric_columns = ['Body', 'NumDoors', 'Drive Type', 'Engine - Liter_Display', 
                               'Engine - CC', 'Engine - Block Type', 'Engine - Cylinders', 
                               'Fuel Type Name', 'Cylinder Type Name', 'Aspiration', 'Engine - CID'
                               ]
            df1[numeric_columns] = df1[numeric_columns].astype(object)

            # Check the data types of the DataFrame
            print(df1.dtypes)

        print(df1.head(5))
    
    except ValueError:
        # Display an error message using the messagebox
        messagebox.showerror("Error", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"No such file as {file_path}")
        return None
    
def Read_data_2():
    """If the file selected is valid this will load the file into the Treeview"""
    global df2  # Declare df as global to update it
    file_path = label_file2["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df2 = pd.read_csv(excel_filename)
        else:
            df2 = pd.read_excel(excel_filename)

            # Convert the numeric columns to str
            numeric_columns = ['Engine - CC', 'Engine - CID', 'Engine - Cylinders', 'NumDoors']
            df2[numeric_columns] = df2[numeric_columns].apply(pd.to_numeric, errors='coerce')

            print(df2.dtypes)

            # Remove leading and trailing whitespaces in df2
            df2 = df2.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        
        print(df2.head(5))
        print(df2.dtypes)
    
    except ValueError:
        # Display an error message using the messagebox
        messagebox.showerror("Error", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"No such file as {file_path}")
        return None
    
  
# Function to check compatibility
# def Check_Compatibility():
#     global df1, df2
#     if df1 is None or df1.empty or df2 is None or df2.empty:
#         messagebox.showerror("Error", "Both dataframes must be loaded before checking compatibility.")
#         return

#     common_columns = set(df1.columns) & set(df2.columns)

#     if not common_columns:
#         messagebox.showerror("Error", "No common columns found between the two dataframes.")
#         return

#     incompatible_columns = []

#     for column in common_columns:
#         values1 = df1[column].astype(str)
#         values2 = df2[column].astype(str)

#         # Filter out empty values in df1
#         values1 = values1[values1 != '']

#         # Check if any non-empty values in df1 are NOT in df2 (case-insensitive)
#         incompatible_values = values1[~values1.str.lower().isin(values2.str.lower())]

#         if not incompatible_values.empty:
#             incompatible_columns.append(column)

#     if incompatible_columns:
#         messagebox.showerror("Error", f"The following columns are not compatible: {', '.join(incompatible_columns)}")
#     else:
#         messagebox.showinfo("Success", "Dataframes are compatible.")







def Run_Matching_Filter(df1, df2):
    if df1 is not None and not df1.empty and df2 is not None and not df2.empty:
        filtered_dfs = []  # List to store filtered DataFrames

        for index, row in df1.iterrows():
            year_beg = row['Year Beg']
            year_end = row['Year End']

            # Custom range filter for the "Year" column
            range_filter = (df2['Year'] >= year_beg) & (df2['Year'] <= year_end)

            # Initialize a boolean mask for the custom column filters
            column_filters = pd.Series(True, index=df2.index)

            # Iterate through the specified columns in df1
            columns_to_filter = [
                'Make', 'Model', 'Submodel', 'Body', 'NumDoors',
                'Drive Type', 'Engine - Liter_Display', 'Engine - CC',
                'Engine - Block Type', 'Engine - Cylinders', 'Fuel Type Name',
                'Cylinder Type Name', 'Aspiration', 'Engine - CID'
            ]

            for column in columns_to_filter:
                value = row[column]
                if pd.notna(value):
                    # Custom column filter
                    column_filters &= df2[column].astype(str).str.contains(str(value), case=False, na=False)

            # Combine the custom range filter and column filters
            combined_filter = range_filter & column_filters

            # Filtered rows in df2 for this row in df1
            filtered_rows = df2.loc[combined_filter].copy()  # Use .copy() to avoid SettingWithCopyWarning

            # Add "PartNumber" and "Notes" columns for each filtered row
            filtered_rows['PartNumber'] = row.get('PartNumber', '')
            filtered_rows['Notes'] = row.get('Notes', '')

            # Append the filtered DataFrame to the list
            filtered_dfs.append(filtered_rows)

        # Concatenate all filtered DataFrames to produce the final result
        filtered_df2 = pd.concat(filtered_dfs)

        # Check for and remove duplicates based on all columns
        filtered_df2 = filtered_df2.drop_duplicates()

        # Generate the current date and time as a string
        current_datetime = datetime.now().strftime("%Y-%m-%d")

        # Define the output file name with the date and time
        output_file_name = f"RunFilter_Output_{current_datetime}.xlsx"

        # Export the filtered data to a new Excel file
        file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
        filtered_df2.to_excel(file_path, index=False, freeze_panes=(1, 1))
        
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

        print(filtered_df2)
        print(filtered_df2.dtypes)
    else:
        print("No data to display")

    return filtered_df2

root.mainloop()








 