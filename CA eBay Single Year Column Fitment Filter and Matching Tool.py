
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
from tkinter import messagebox
import os
import xlwings as xw
from datetime import datetime
import datetime
import customtkinter
from collections import defaultdict
import re


# Declare df1 and df2 as global variables
df1 = None
df2 = None
df3 = None

# Declare global variables for file paths
file_path_1 = None
file_path_2 = None

# Setting up theme of the app
customtkinter.set_appearance_mode("system")

# Setting up them of your components
customtkinter.set_default_color_theme("blue")

# initalise the tkinter GUI
root = customtkinter.CTk()
root.title("CA eBay Single Year Column Fitment Filter and Matching Tool")

root.geometry("900x650") # set the root dimensions
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.

frame1 = tk.LabelFrame(root, text="Functions", bg="lightgrey", fg="black", font=("Arial", 12, "bold"))
frame1.place(height=150, width=800, relx=0.01)

file_frame = tk.LabelFrame(root, text="Display File", bg="lightgrey", fg="black", font=("Arial", 12, "bold"))
file_frame.place(height=200, width=800, rely=0.65, relx=0.01)

file_frame1 = tk.LabelFrame(root, text="Display File", bg="lightgrey", fg="black", font=("Arial", 12, "bold"))
file_frame1.place(height=200, width=800, rely=0.30, relx=0.01)


# Buttons
button1 = customtkinter.CTkButton(file_frame, text="Select Custom Fitment File", command=lambda: file_dialog_1(), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button1.place(rely=0.2, relx=0.01)

button2 = customtkinter.CTkButton(file_frame, text="Read File and Save File Path Link", command=lambda: read_data_1(), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button2.place(rely=0.4, relx=0.01)

button3 = customtkinter.CTkButton(file_frame1, text="Select eBay MVL File", command=lambda: file_dialog_2(), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button3.place(rely=0.2, relx=0.01)

button4 = customtkinter.CTkButton(file_frame1, text="Read File and Save File Path Link", command=lambda: read_data_2(), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button4.place(rely=0.4, relx=0.01)

button5 = customtkinter.CTkButton(frame1, text="Run Matching Filter", command=lambda: run_matching_filter(df1, df2), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button5.place(rely=0.2, relx=0.01)

button6 = customtkinter.CTkButton(frame1, text="Check Valid Values", command=lambda: print_compatibility(), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button6.place(rely=0.5, relx=0.01)

button7 = customtkinter.CTkButton(frame1, text="Create CA eBay Compatibilty", command=lambda: read_data_3(), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button7.place(rely=0.2, relx=0.3)

button8 = customtkinter.CTkButton(frame1, text="Filter eBay MVL File", command=lambda: pre_filter_eBay_MVL_File(df1, df2), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button8.place(rely=0.5, relx=0.3)

button9 = customtkinter.CTkButton(frame1, text="Run Matching Filter_Fits All", command=lambda: run_matching_filter_2(df1, df2), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button9.place(rely=0.2, relx=0.65)

button10 = customtkinter.CTkButton(frame1, text="Create CA eBay Compatibility_Fits All", command=lambda: read_data_4(), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button10.place(rely=0.5, relx=0.65)

label_file1 = ttk.Label(file_frame, text="No File Selected", background="lightgrey", foreground="blue", font=("Arial", 11, "bold"))
label_file1.place(rely=0, relx=0)

label_file2 = ttk.Label(file_frame1, text="No File Selected", background="lightgrey", foreground="blue", font=("Arial", 11, "bold"))
label_file2.place(rely=0, relx=0)


def file_dialog_1():
    """This Function will open the file explorer and assign the chosen file path to label_file1"""
    global file_path_1  # Declare as global to update it
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file1["text"] = filename
    file_path_1 = filename  # Update the global variable
    return None

def file_dialog_2():
    """This Function will open the file explorer and assign the chosen file path to label_file2"""
    global file_path_2  # Declare as global to update it
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file2["text"] = filename
    file_path_2 = filename  # Update the global variable
    return None

def read_data_1():
    """If the file selected is valid this will load the file into the Treeview"""
    global df1  # Declare df as global to update it
    file_path = file_path_1
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df1 = pd.read_csv(excel_filename)
        else:
            df1 = pd.read_excel(excel_filename)
           
           
            # Convert only specific columns to object
            columns_to_convert_to_object = ['Make', 'Model', 'Submodel',
                'Body', 'Drive Type', 'Engine - Block Type', 'Engine - Liter_Display',
                'Fuel Type Name', 'Cylinder Type Name', 'Aspiration'
            ]
            df1[columns_to_convert_to_object] = df1[columns_to_convert_to_object].astype(object)
            
            # Convert the [NumDoors] data type into float
            df1['NumDoors'] = df1['NumDoors'].astype(float)

            # Convert the [Year] data type into Int64
            df1['Year'] = pd.to_numeric(df1['Year'], errors='coerce')
            df1['Year'] = df1['Year'].astype('Int64')

            print(df1.dtypes)


        # Show "Complete" message when the function is done
        messagebox.showinfo("Read and Save", "completed successfully!")
        
    # Handling error
    except ValueError:
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"No such file as {file_path}")
        return None


def read_data_2():
    """If the file selected is valid this will load the file into the Treeview"""
    global df2  # Declare df as global to update it
    file_path = file_path_2
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
        # Show "Complete" message when the function is done
        messagebox.showinfo("Read and Save", "completed successfully!")
    
    # Handling error
    except ValueError:
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"No such file as {file_path}")
        return None
    
def read_data_3():
    """If the file selected is valid this will load the file into the Treeview"""
    global df3  # Declare df as global to update it
    
    # Specify the path to your local directory where you downloaded files
    local_directory = os.path.expanduser("~/Desktop")

    # List all files in the local directory
    files_in_directory = os.listdir(local_directory)

    # Define the substring you want to filter for
    substring = "RunFilter_Output"

    # Filter files to include only CSV files that contain the specified substring
    excel_files = [file for file in files_in_directory if file.endswith(".xlsx") and substring in file]

    # If you want to find the latest file based on modification time, you can use this:
    latest_file = max(excel_files, key=lambda x: os.path.getmtime(os.path.join(local_directory, x)))

    # Construct the full path to the latest file
    latest_file_path = os.path.join(local_directory, latest_file)

    try:
        if latest_file_path[-4:] == ".csv":
            df3 = pd.read_csv(latest_file_path)
        else:
            df3 = pd.read_excel(latest_file_path)

            # Initialize an empty dictionary to store the formatted strings
            formatted_data = {}

            # Iterate through each row in the DataFrame
            for index, row in df3.iterrows():
                # Format the data in each row
                formatted_row = f"{row['Year']}|{row['Make']}|{row['Model']}|{row['Trim']}|{row['Engine']}::{row['Notes']}"

                # If this 'PartNumber' is not in the dictionary, add it
                if row['PartNumber'] not in formatted_data:
                    formatted_data[row['PartNumber']] = formatted_row
                else:  # If this 'PartNumber' is already in the dictionary, append the new data
                    formatted_data[row['PartNumber']] += '^^' + formatted_row

            # Sort the dictionary by key (i.e., 'PartNumber') in ascending order
            formatted_data = dict(sorted(formatted_data.items()))

            # Convert the dictionary to a list of strings
            final_text_list = [f"{part_number}\tUNSHIPPED\t{data}" for part_number, data in formatted_data.items()]

            # Join the list of strings with '\n' as the delimiter to create the final text
            final_text = '\n'.join(final_text_list)

            # Add a header to the final text
            final_text = "Inventory Number\tQuantity Update Type\ta2Listing Fitment\n" + final_text

            final_text = final_text.replace("nan", "")

            # Print the final text
            print(final_text)

            # Show "Complete" message when the function is done
            messagebox.showinfo("CA eBay Compatibility", "Compatibilty Complete!")

            # Define the output file name
            output_file_name = "FitmentOutput.txt"

            # Define the full path to the output text file on your Desktop
            output_file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)

            # Write the final text to the output text file
            with open(output_file_path, 'w') as txt_file:
                txt_file.write(final_text)

    # Handling error
    except ValueError:
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"No such file as {latest_file_path}")
        return None
    
def read_data_4():
    global df3  # Declare df as global to update it
    
    # Specify the path to your local directory where you downloaded files
    local_directory = os.path.expanduser("~/Desktop")

    # List all files in the local directory
    files_in_directory = os.listdir(local_directory)

    # Filter for Excel files containing the specified substring
    substring = "RunFilter_Output"
    excel_files = [file for file in files_in_directory if file.endswith(".xlsx") and substring in file]

    # Get the latest file based on modification time
    latest_file = max(excel_files, key=lambda x: os.path.getmtime(os.path.join(local_directory, x)))
    latest_file_path = os.path.join(local_directory, latest_file)

    try:
        if latest_file_path.endswith(".csv"):
            df3 = pd.read_csv(latest_file_path)
        else:
            df3 = pd.read_excel(latest_file_path)

        # Filter data based on specified conditions
        years_by_part_make_model = {}
        years_by_part_make_model_trim = {}
        
        for _, row in df3.iterrows():
            if pd.notna(row['Year']) and pd.notna(row['Make']) and pd.notna(row['Model']) and pd.notna(row['PartNumber']) and all(pd.isnull(row[col]) for col in [
                'ePID', 'Aspiration', 'Body', 'Cylinder Type Name', 
                'DisplayName', 'Drive Type', 'Engine',
                'Engine - Block Type', 'Engine - CC',
                'Engine - CID', 'Engine - Cylinders',
                'Engine - Liter_Display', 'Fuel Type Name',
                'KBB_MODEL', 'NumDoors', 'Parts Model', 'Submodel',
                'Trim'
            ]):
                combined_key = (row['PartNumber'], row['Make'], row['Model'])
                years_by_part_make_model.setdefault(combined_key, []).append(row['Year'])

            elif pd.notna(row['Year']) and pd.notna(row['Make']) and pd.notna(row['Model']) and pd.notna(row['Trim']) and pd.notna(row['PartNumber']) and all(pd.isnull(row[col]) for col in [
                'ePID', 'Aspiration', 'Body', 'Cylinder Type Name', 
                'DisplayName', 'Drive Type', 'Engine',
                'Engine - Block Type', 'Engine - CC',
                'Engine - CID', 'Engine - Cylinders',
                'Engine - Liter_Display', 'Fuel Type Name',
                'KBB_MODEL', 'NumDoors', 'Parts Model', 'Submodel'
                
            ]):
                combined_key = (row['PartNumber'], row['Make'], row['Model'], row['Trim'])
                years_by_part_make_model_trim.setdefault(combined_key, []).append(row['Year'])

        # Iterate the year in the dict in years_by_part_make_model
        for key, years in years_by_part_make_model.items():
           if len(set(years)) == 1:  # Check if all years are the same
              year = years[0]
              years_by_part_make_model[key] = year
           else:
              year_range = f"{min(years)}-{max(years)}" if years else "No valid years found"
              years_by_part_make_model[key] = year_range

        # Iterate the year in the dict in years_by_part_make_model_trim
        for key, years in years_by_part_make_model_trim.items():
           if len(set(years)) == 1:  # Check if all years are the same
              year = years[0]
              years_by_part_make_model_trim[key] = year
           else:
              year_range = f"{min(years)}-{max(years)}" if years else "No valid years found"
              years_by_part_make_model_trim[key] = year_range
        
        formatted_data = defaultdict(str)

        for _, row in df3.iterrows():
            if pd.notna(row['Year']) and pd.notna(row['Make']) and pd.notna(row['Model']) and pd.notna(row['PartNumber']) and all(pd.isnull(row[col]) for col in [
                'ePID', 'Aspiration', 'Body', 'Cylinder Type Name', 
                'DisplayName', 'Drive Type', 'Engine',
                'Engine - Block Type', 'Engine - CC',
                'Engine - CID', 'Engine - Cylinders',
                'Engine - Liter_Display', 'Fuel Type Name',
                'KBB_MODEL', 'NumDoors', 'Parts Model', 'Submodel',
                'Trim'
            ]): 
                
                combined_key = (row['PartNumber'], row['Make'], row['Model'])
                formatted_string = f"{years_by_part_make_model.get(combined_key, 'No valid years found')}|{row['Make']}|{row['Model']}::{row['Notes']}"
                key = f"{row['PartNumber']}_{row['Make']}_{row['Model']}"
                formatted_data[key] = formatted_string
                
            
            elif pd.notna(row['Year']) and pd.notna(row['Make']) and pd.notna(row['Model']) and pd.notna(row['Trim']) and pd.notna(row['PartNumber']) and all(pd.isnull(row[col]) for col in [
                'ePID', 'Aspiration', 'Body', 'Cylinder Type Name', 
                'DisplayName', 'Drive Type', 'Engine',
                'Engine - Block Type', 'Engine - CC',
                'Engine - CID', 'Engine - Cylinders', 
                'Engine - Liter_Display', 'Fuel Type Name',
                'KBB_MODEL', 'NumDoors', 'Parts Model', 'Submodel'              
            ]): 
                combined_key = (row['PartNumber'], row['Make'], row['Model'], row['Trim'])
                formatted_string = f"{years_by_part_make_model_trim.get(combined_key, 'No valid years found')}|{row['Make']}|{row['Model']}|{row['Trim']}::{row['Notes']}"
                key = f"{row['PartNumber']}_{row['Make']}_{row['Model']}_{row['Trim']}"
                formatted_data[key] = formatted_string
                          
            else:
                formatted_row = f"{row['Year']}|{row['Make']}|{row['Model']}|{row['Trim']}|{row['Engine']}::{row['Notes']}"
                key = f"{row['PartNumber']}_{row['Make']}_{row['Model']}"
                formatted_data[key] += '^^' + formatted_row
                                                                                
        # Sort the dictionary by key (i.e., 'PartNumber') in ascending order
        formatted_data = dict(sorted(formatted_data.items(), key=lambda item: item[1], reverse=False))

        final_text_list = []

        # Group the fitment data by Inventory Number
        fitment_by_inventory_number = {}
        for part_number, data in formatted_data.items():
            if isinstance(part_number, str):
               inventory_number = part_number.split('_')[0]
               fitment_by_inventory_number.setdefault(inventory_number, []).append(data)

        # Create the final text list with fitment data grouped by Inventory Number
        for inventory_number, fitments in fitment_by_inventory_number.items():
            fitment_string = '^^'.join(fitments)
            final_text_list.append(f"{inventory_number}\tUNSHIPPED\t{fitment_string}")

        # Define a function to extract the number from the 'Inventory Number'
        def get_inventory_number(s):
           return int(s.split('\t')[0].split('-')[1])
        
        # Sort the list based on the inventory number
        final_text_list.sort(key=get_inventory_number)

        # Remove the element '^^^^' and replace with '^^'
        final_text_list = [s.replace('^^^^', '^^') for s in final_text_list]
        

        # Remove the element '^^' at the beginning of string text fitment
        final_text_list = [re.sub(r'UNSHIPPED\s\^\^', 'UNSHIPPED\t', s) for s in final_text_list] 

        # Join the list of strings with '\n' as the delimiter to create the final text
        final_text = "Inventory Number\tQuantity Update Type\ta2Listing Fitment\n" + '\n'.join(final_text_list)

        # Suppose final_text and final_text_list are defined earlier in your code
        final_text = final_text.replace("::nan", "")
            
        # Show "Complete" message when the function is done
        messagebox.showinfo("CA eBay Compatibility", "Compatibility Complete")
        
        # Define the output file name
        output_file_name = "FitmentOutput.txt"

        # Define the full path to the output text file on your Desktop
        output_file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
        
        # Write the final text to the output text file
        with open(output_file_path, 'w') as txt_file:
            txt_file.write(final_text)

    except ValueError:
        messagebox.showerror("Error", f"The file you have chosen is invalid")
    except FileNotFoundError:
        messagebox.showerror("Error", f"No such file as {latest_file_path}")

    
def pre_filter_eBay_MVL_File(df1, df2):
    if df1 is not None and not df1.empty and df2 is not None and not df2.empty:
        filtered_dfs = []  # List to store filtered DataFrames
       
        for _, row in df1.iterrows():
            # Initialize a boolean mask for the custom column filters
            column_filters = pd.Series(True, index=df2.index)

            # Iterate through the specified columns in df1
            columns_to_filter = ['Year', 'Make', 'Model']

            for column in columns_to_filter:
                value = row[column]
                if pd.notna(value):
                    # Custom column filter
                    column_filters &= (df2[column].astype(str) == str(value))
                    
            # Filtered rows in df2 for this row in df1
            filtered_rows = df2.loc[column_filters].copy()  # Use .copy() to avoid SettingWithCopyWarning

            # Append the filtered DataFrame to the list
            filtered_dfs.append(filtered_rows)

        # Concatenate all filtered DataFrames to produce the final result
        filtered_df3 = pd.concat(filtered_dfs)

        # Check for and remove duplicates based on all columns
        filtered_df3 = filtered_df3.drop_duplicates()

        # Generate the current date and time as a string
        current_datetime = datetime.now().strftime("%Y-%m-%d")

        # Define the output file name with the date and time
        output_file_name = f"FilterMVLFile_Output_{current_datetime}.xlsx"

        # Show "Complete" message when the function is done
        messagebox.showinfo("MVL File Output", "Filter Complete")

        # Export the filtered data to a new Excel file
        file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
        filtered_df3.to_excel(file_path, index=False, freeze_panes=(1, 0))
        
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

        print(filtered_df3)
        print(filtered_df3.dtypes)
    else:
        print("No data to display")

    return filtered_df3


def Check_Compatibility(df1, df2):
    """Checks the compatibility of two dataframes, ignoring empty values in df1.

    Args:
        df1 (pandas.DataFrame): The first dataframe.
        df2 (pandas.DataFrame): The second dataframe.

    Returns:
        pandas.DataFrame: A DataFrame with details of incompatibility if the dataframes are not compatible, None otherwise.
    """
    
    try:
        # Check if both dataframes are loaded and not empty.
        if df1 is None or df2 is None:
            raise ValueError("Both dataframes must be provided.")

        # Find the common columns between the two dataframes.
        common_columns = set(df1.columns) & set(df2.columns)

        # Check if there are any common columns.
        if not common_columns:
            return pd.DataFrame([["No common columns found between the two dataframes.", None, None]], columns=['Column', 'Index', 'Value'])
        
        # Initialize a list to store messages about incompatible columns
        incompatible_columns = []

        # Check if all values in the common columns of df1 (ignoring empty values) are present in df2.
        for column in common_columns:
            values1 = df1[column].astype(str).replace('nan', '').str.strip()
            values2 = df2[column].astype(str).replace('nan', '').str.strip()

            # Filter out empty values in df1
            non_empty_values1 = values1[values1 != '']

            # Check if any non-empty value in df1 is not in df2.
            if not non_empty_values1.isin(values2).all():
                incompatible_values = non_empty_values1[~non_empty_values1.isin(values2)]
                for index, value in incompatible_values.items():
                    incompatible_columns.append([column, index, value])
        
        # If we found any incompatible columns, return a DataFrame with their details
        if incompatible_columns:
            return pd.DataFrame(incompatible_columns, columns=['Column', 'Index', 'Value'])
        
        # If we reach this point, all non-empty values in the common columns of df1 are present in df2, so the dataframes are compatible.
        return pd.DataFrame([["Dataframes are compatible.","", ""]], columns=['Column', 'Index', 'Value'])
    
    except ValueError as e:
        messagebox.showinfo("Error", str(e))

def print_compatibility():
    compatibility = Check_Compatibility(df1, df2)  # replace df1 and df2 with your actual dataframes
    if compatibility is not None:
        if isinstance(compatibility, str):
            messagebox.showinfo("Compatibility Check", compatibility)
            save_output_to_csv(compatibility, 'compatibility.csv')
        else:
            compatibility_str = "\n".join(compatibility['Column'].astype(str) + "\t" + compatibility['Index'].astype(str) + "\t" + compatibility['Value'].astype(str))
            messagebox.showinfo("Compatibility Check", compatibility_str)
            save_output_to_csv(compatibility_str, 'error_report.csv')

def save_output_to_csv(error_report, filename):
    # Split the output string into lines
    lines = error_report.split('\n')

    # Split each line into columns and store them in a list
    data = [line.split('\t') for line in lines]

    # Convert the list into a DataFrame
    df = pd.DataFrame(data, columns=['Column', 'Index', 'Value'])

    # Save the DataFrame to a CSV file
    file_path = os.path.join(os.path.expanduser("~"), "Desktop", filename)
    df.to_csv(file_path, index=False, header=True)

def run_matching_filter(df1, df2):
    if df1 is not None and not df1.empty and df2 is not None and not df2.empty:
        filtered_dfs = []  # List to store filtered DataFrames

        for _, row in df1.iterrows():
            # Initialize a boolean mask for the custom column filters
            column_filters = pd.Series(True, index=df2.index)

            # Iterate through the specified columns in df1
            columns_to_filter = [
                'Year', 'Make', 'Model', 'Submodel', 'Body', 'NumDoors',
                'Drive Type', 'Engine - Liter_Display', 'Engine - CC',
                'Engine - Block Type', 'Engine - Cylinders', 'Fuel Type Name',
                'Cylinder Type Name', 'Aspiration', 'Engine - CID'
            ]

            for column in columns_to_filter:
                value = row[column]
                if pd.notna(value):
                    # Custom column filter
                    column_filters &= (df2[column].astype(str) == str(value))

            # Filtered rows in df2 for this row in df1
            filtered_rows = df2.loc[column_filters].copy()  # Use .copy() to avoid SettingWithCopyWarning

            # Add "PartNumber" and "Notes" columns for each filtered row
            filtered_rows['PartNumber'] = row.get('PartNumber', '')
            filtered_rows['Notes'] = row.get('Notes', '')

            # Append the filtered DataFrame to the list
            filtered_dfs.append(filtered_rows)

        # Concatenate all filtered DataFrames to produce the final result
        filtered_df2 = pd.concat(filtered_dfs)

        # Add this line to sort 'Year' column in descending order
        filtered_df2 = filtered_df2.sort_values(['PartNumber', 'Year'], ascending=[False, False])

        # Check for and remove duplicates based on all columns
        filtered_df2 = filtered_df2.drop_duplicates()

        # Generate the current date and time as a string
        current_datetime = datetime.datetime.now().strftime("%Y-%m-%d")

        # Define the output file name with the date and time
        output_file_name = f"RunFilter_Output_{current_datetime}.xlsx"

        # Show "Complete" message when the function is done
        messagebox.showinfo("Run Matching Filter", "Filter Complete!")

        # Export the filtered data to a new Excel file
        file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
        filtered_df2.to_excel(file_path, index=False, freeze_panes=(1, 0))
        
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


def run_matching_filter_2(df1, df2):
    if df1 is not None and not df1.empty and df2 is not None and not df2.empty:
                                 
        # Initialize a list to store filtered DataFrames
        filtered_dfs = []
        
        # Define the special condition as a separate function
        def is_special_condition(row):
            return (
                pd.notna(row['PartNumber']) and
                pd.notna(row['Year']) and pd.notna(row['Make']) and pd.notna(row['Model']) and
                all(pd.isnull(row[col]) for col in [
                    'Submodel', 'Body', 'NumDoors', 'Drive Type',
                    'Engine - Liter_Display', 'Engine - CC',
                    'Engine - Block Type', 'Engine - Cylinders',
                    'Fuel Type Name', 'Cylinder Type Name',
                    'Aspiration', 'Engine - CID']
                )
            )

        # Apply the special condition to each row in df1
        special_condition_rows = df1.apply(is_special_condition, axis=1)

        # Separate the rows that meet the special condition from the rest of the data
        special_condition_data = df1[special_condition_rows]
        df1 = df1[~special_condition_rows]

        # Process the special_condition_data by selecting only the necessary columns
        special_condition_data = special_condition_data[['PartNumber', 'Year', 'Make', 'Model', 'Submodel', 'Body', 
                                                         'NumDoors', 'Drive Type', 'Engine - Liter_Display',
                                                         'Engine - CC', 'Engine - Block Type', 'Engine - Cylinders', 
                                                         'Fuel Type Name', 'Cylinder Type Name', 'Aspiration', 'Engine - CID', 'Notes' ]]

        for _, row in special_condition_data.iterrows():

            # Initialize a boolean mask for the custom column filters
            column_filters = pd.Series(True, index=df2.index)

            # Iterate through the specified columns in df1
            columns_to_filter = [
                'Year', 'Make', 'Model', 'Submodel', 'Body', 'NumDoors',
                'Drive Type', 'Engine - Liter_Display', 'Engine - CC',
                'Engine - Block Type', 'Engine - Cylinders',
                'Fuel Type Name', 'Cylinder Type Name', 'Aspiration', 'Engine - CID'
            ]

            for column in columns_to_filter:
                value = row[column]
                if pd.notna(value):
                    # Custom column filter
                    column_filters &= (df2[column].astype(str) == str(value))


            # Filtered rows in df2 for this row in df1
            filtered_rows = df2.loc[column_filters].copy()

            # Add "PartNumber" and "Notes" columns for each filtered row
            filtered_rows['PartNumber'] = row.get('PartNumber', '')
            filtered_rows['Notes'] = row.get('Notes', '')

            # Append the filtered DataFrame to the list
            filtered_dfs.append(filtered_rows)
               
        if filtered_dfs:
            filtered_df = pd.concat(filtered_dfs)
        else:
            filtered_df = pd.DataFrame()

        # # Check if there are no objects to concatenate
        # if not filtered_dfs:
        #    messagebox.showinfo("Error", "No Fits All, Use Run Matching Filter")
        #    return

        # Concatenate all filtered DataFrames to produce the final result        
        #filtered_df = pd.concat(filtered_dfs)
       
        # List of columns to retain
        columns_to_retain = ['Year', 'Model', 'Make', 'PartNumber', 'Notes']

        # Iterate through the DataFrame and set empty values for other columns
        for column in filtered_df.columns:
            if column not in columns_to_retain:
                filtered_df[column] = ''

        # Add this line to sort 'Year' column in descending order
        # filtered_df = filtered_df.sort_values(['PartNumber', 'Year'], ascending=[False, False])
        
        # Add this line to check if 'PartNumber' column exists before sorting
        if 'PartNumber' in filtered_df.columns:
            # Add this line to sort 'Year' column in descending order
            filtered_df = filtered_df.sort_values(['PartNumber', 'Year'], ascending=[False, False])


        # Define a function to process the remaining data
        def process_remaining_data(remaining_data, df2):
            filtered_dfs2 = []

            for _, row in remaining_data.iterrows():

                # Initialize a boolean mask for the custom column filters
                column_filters_2 = pd.Series(True, index=df2.index)

                # Check if any of the combinations are present in df1
                combinations = [
                   ['Make', 'Model', 'Submodel'],
                   ['Make', 'Model', 'Body'],
                   ['Make', 'Model', 'NumDoors'],
                   ['Make', 'Model', 'Submodel', 'Body', 'NumDoors']
                ]
                combination_present = any(all(pd.notna(row[c]) for c in combo) for combo in combinations)

                # Check if any other fields are present
                other_fields = ['Drive Type', 'Engine - Liter_Display', 'Engine - CC',
                                'Engine - Block Type', 'Engine - Cylinders',
                                'Fuel Type Name', 'Cylinder Type Name', 'Aspiration', 'Engine - CID']
                other_fields_present = any(pd.notna(row[c]) for c in other_fields)


                # If any of the combinations are present, apply columns_to_retain_2
                if combination_present:
                    columns_to_filter_2 = ['Make', 'Model', 'Submodel', 'Body', 'NumDoors']
                else:
                    columns_to_filter_2 = ['Make', 'Model', 'Submodel', 'Body', 'NumDoors',
                                           'Drive Type', 'Engine - Liter_Display', 'Engine - CC',
                                           'Engine - Block Type', 'Engine - Cylinders',
                                           'Fuel Type Name', 'Cylinder Type Name', 'Aspiration', 'Engine - CID']

          
                for column in columns_to_filter_2:
                    value = row[column]
                    if pd.notna(value):
                        # Custom column filter
                        column_filters_2 &= (df2[column].astype(str) == str(value))

                # Filtered rows in df2 for this row in df1
                filtered_rows_2 = df2.loc[column_filters_2].copy()

                # Add "PartNumber" and "Notes" columns for each filtered row
                filtered_rows_2['PartNumber'] = row.get('PartNumber', '')
                filtered_rows_2['Notes'] = row.get('Notes', '')

                # If any of the combinations are present and no other fields are present, apply columns_to_retain_2
                if combination_present and not other_fields_present:

                    # List of columns to retain
                   columns_to_retain_2 = ['Trim', 'Year', 'Model', 'Make', 'PartNumber', 'Notes']

                   # Iterate through the DataFrame and set empty values for other columns
                   for column in filtered_rows_2.columns:
                       if column not in columns_to_retain_2:
                          filtered_rows_2[column] = ''
     
                # Append the filtered DataFrame to the list
                filtered_dfs2.append(filtered_rows_2)

            # Concatenate all filtered DataFrames for remaining_data
            if filtered_dfs2:  # Check if filtered_dfs2 is not empty
               filtered_dfs2 = pd.concat(filtered_dfs2)
            else:
               filtered_dfs2 = pd.DataFrame() # Return empty DataFrame

            if filtered_df.empty and filtered_dfs2.empty:
               print("Both DataFrames are empty")
               # return or error handling
            elif filtered_df.empty: 
               final_result = filtered_dfs2 
            elif filtered_dfs2.empty:
               final_result = filtered_df
            else:
               final_result = pd.concat([filtered_df, filtered_dfs2])
            
            # Generate the current date and time as a string
            current_datetime = datetime.datetime.now().strftime("%Y-%m-%d")

            # Define the output file name with the date and time
            output_file_name = f"RunFilter_Output_{current_datetime}.xlsx"

            # Show "Complete" message when the function is done
            messagebox.showinfo("Run Matching Filter", "Filter Complete")

            # Add this line to sort 'Year' column in descending order
            final_result = final_result.sort_values(by=['PartNumber', 'Year'], ascending=[True, False])

            # Export the filtered data to a new Excel file
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
            final_result.to_excel(file_path, index=False, freeze_panes=(1, 0))

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

            print(final_result)
            print(final_result.dtypes)
        
        # Process the remaining data with the function
        process_remaining_data(df1, df2)        
    
    else:
        print("No Fits All, User Run Matching Filter")
        return

root.mainloop()

