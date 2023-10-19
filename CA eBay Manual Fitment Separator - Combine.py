import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
from tkinter import messagebox
import os
import re
import xlwings as xw
import customtkinter as ctk

# Declare df as a global variable
df = None
df1 = None
df2 = None


# Declare global variables for file paths
file_path = None
file_path_1 = None
file_path_2 = None

# Setting up theme of the app
ctk.set_appearance_mode("system")

# Setting up them of your components
ctk.set_default_color_theme("green")

# Initialize the tkinter GUI
root = ctk.CTk()
root.title("CA eBay Manual Fitment Separator - Combine")

root.geometry("650x650")  # set the root dimensions
root.pack_propagate(False)  # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0)  # makes the root window fixed in size.

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File", bg="lightgrey", fg="black", font=("Arial", 15, "bold"))
file_frame.place(height=200, width=800, rely=0, relx=0.01)

# Frame for open file dialog
file_frame2 = tk.LabelFrame(root, text="Display File", bg="lightgrey", fg="black", font=("Arial", 15, "bold"))
file_frame2.place(height=200, width=800, rely=0.30, relx=0.01)

file_frame3 = tk.LabelFrame(root, text="Display File", bg="lightgrey", fg="black", font=("Arial", 15, "bold"))
file_frame3.place(height=200, width=800, rely=0.65, relx=0.01)


# Buttons
button1 = ctk.CTkButton(file_frame, text="Browse For File", command=lambda: File_dialog(), fg_color="white",
                                  text_color='black', font=('Arial', 15, 'bold'))
button1.place(rely=0.2, relx=0.01)

button2 = ctk.CTkButton(file_frame, text="Run Separator", command=lambda: Transform_1(), fg_color='white',
                                  text_color='black', font=('Arial', 15, 'bold'))
button2.place(rely=0.45, relx=0.01)

button3 = ctk.CTkButton(file_frame, text="Run Combine", command=lambda: Transform_2(), fg_color='white',
                                  text_color='black', font=('Arial', 15, 'bold'))
button3.place(rely=0.7, relx=0.01)

button4 = ctk.CTkButton(file_frame2, text="Select eBay MVL File", command=lambda: File_dialog_2(), fg_color="light green",
                                  text_color='black', font=('Arial', 15, 'bold'))
button4.place(rely=0.2, relx=0.01)

button5 = ctk.CTkButton(file_frame3, text="Select Fitment File", command=lambda: File_dialog_3(), fg_color='light green',
                                  text_color='black', font=('Arial', 15, 'bold'))
button5.place(rely=0.2, relx=0.01)

button6 = ctk.CTkButton(file_frame2, text="Read File and Save File Path Link", command=lambda: Read_data_1(), fg_color='light green', 
                                     text_color='black', font=('Arial', 15, 'bold'))
button6.place(rely=0.5, relx=0.01)

button7 = ctk.CTkButton(file_frame3, text="Read File and Save File Path Link", command=lambda: Read_data_2(), fg_color='light green', 
                                     text_color='black', font=('Arial', 15, 'bold'))
button7.place(rely=0.5, relx=0.01)

button8 = ctk.CTkButton(file_frame, text="Check Valid Values", command=lambda: print_compatibility(), fg_color='light green', 
                                    text_color='black', font=('Arial', 15, 'bold'))
button8.place(rely=0.2, relx=0.3)

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected", background="lightgrey", foreground="blue",
                       font=("Arial", 9, "bold"))
label_file.place(rely=0, relx=0)

label_file2 = ttk.Label(file_frame2, text="No File Selected", background="lightgrey", foreground="blue",
                       font=("Arial", 9, "bold"))
label_file2.place(rely=0, relx=0)

label_file3 = ttk.Label(file_frame3, text="No File Selected", background="lightgrey", foreground="blue",
                       font=("Arial", 9, "bold"))
label_file3.place(rely=0, relx=0)

def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file1"""
    global file_path  # Declare as global to update it
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file["text"] = filename
    file_path = filename  # Update the global variable
    return None

def File_dialog_2():
    """This Function will open the file explorer and assign the chosen file path to label_file2"""
    global file_path_1  # Declare as global to update it
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file2["text"] = filename  # Update the label in file_frame2
    file_path_1 = filename  # Update the global variable
    return None

def File_dialog_3():
    """This Function will open the file explorer and assign the chosen file path to label_file3"""
    global file_path_2  # Declare as global to update it
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file3["text"] = filename  # Update the label in file_frame3
    file_path_2 = filename  # Update the global variable
    return None


def Transform_1():
    """If the file selected is valid, this will load the file into the Treeview"""
    global df  # Declare df as global to update it
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

        print(df.dtypes)

        # Initialize an empty list to store the transformed data
        transformed_data = []

        # Regular expression pattern to extract years from year_range
        year_pattern = r'(\d{4})-(\d{4})'

        # Iterate through each row in the DataFrame
        for _, row in df.iterrows():
            # Get the inventory number and fitment from the row
            inventory_number = row['Inventory Number']
            fitment = row['a2Listing Fitment']

            # Split the fitment into multiple fitments if necessary
            fitments = str(fitment).split("^^")          

            # Iterate through each fitment
            for fitment in fitments:
                # Split the fitment into year range, make, model, and notes
                fitment_components = fitment.split("|")

                # Ensure there are at least 2 components for year_range, and fill in missing values
                if len(fitment_components) < 2:
                    fitment_components += [''] * (2 - len(fitment_components))

                # Unpack the components with extra components included in make_model_notes
                year_range, make_model_notes = fitment_components[0], "|".join(fitment_components[1:])
                
                # Extract start and end years using regular expression
                year_match = re.search(year_pattern, year_range)
                if year_match:
                    start_year, end_year = map(int, year_match.groups())
                elif year_range.isdigit():
                    # If the year_range is a single year, use it as both the start and end year
                    start_year = end_year = int(year_range)
                else:
                    # Handle unexpected year_range format
                    start_year, end_year = 0, 0  # Assign default values or handle it as needed

                # Split make_model_notes into make, model, and notes with extra components included in notes
                make_model, notes = make_model_notes.split("::", 1) if "::" in make_model_notes else (make_model_notes, '')
                make, model = make_model.split("|", 1) if "|" in make_model else (make_model, '')
            
                # Iterate through each year in the year range
                for year in range(start_year, end_year + 1):
                    # Append the transformed line to the list of transformed data
                    transformed_data.append([inventory_number, str(year), make, model, notes])

        # Convert list of lists to DataFrame
        final_df = pd.DataFrame(transformed_data, columns=['Inventory Number', 'Year', 'Make', 'Model', 'Notes'])
        
        print(final_df.dtypes)

        # Define the output file name
        output_file_name = "eBay Fitment Separator Final.xlsx"

        # Show "Complete" message when the function is done
        messagebox.showinfo("Separator", "Complete!")

        # Define the full path to the output text file on your Desktop
        output_file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
        final_df.to_excel(output_file_path, index=False, freeze_panes=(1, 0))

        print(final_df)

        # Open the Excel file and set all columns width to 15
        with xw.App(visible=False) as app:
            wb = xw.Book(output_file_path)

            # Loop through all worksheets in the workbook
            for ws in wb.sheets:
                # Loop through all columns in the worksheet
                for column in ws.api.UsedRange.Columns:
                    column.ColumnWidth = 15
                    print(f"Column {column.Column}: Width set to {column.ColumnWidth}")

            # Save the workbook if needed
            wb.save()
            print(f"Workbook saved: {output_file_path}")

            # Close the workbook
            wb.close()
            print("Workbook closed")

    except ValueError as e:
        print("Error:", e)  # Print the specific error message
        tk.messagebox.showerror("Information", f"The file you have chosen is invalid")
        return None
    except FileNotFoundError as e:
        print("Error:", e)  # Print the specific error message
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    

def Transform_2():
    """If the file selected is valid, this will load the file into the Treeview"""
    global df  # Declare df as global to update it
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)
        
        # Fill missing values in 'Inventory Number', 'Year', 'Make', 'Model', 'Notes' with an empty string
        List_columns = ['Inventory Number', 'Year', 'Make', 'Model', 'Notes']
        
        df[List_columns] = df[List_columns].fillna('')

        # Group by 'Inventory Number', 'Make', 'Model', and 'Notes' and aggregate 'Year' into a list
        grouped = df.groupby(['Inventory Number', 'Make', 'Model', 'Notes'])['Year'].apply(list).reset_index()

        # Define a function to convert a list of years into a range string
        def year_range(years):
            min_year = min(years)
            max_year = max(years)
            if min_year == max_year:
                return str(min_year)
            else:
                return f"{min_year}-{max_year}"

        # Apply the year_range function to the 'Year' column
        grouped['Year'] = grouped['Year'].apply(year_range)

        # Combine 'Year', 'Make', 'Model', and 'Notes' into a single string in the 'a2Listing Fitment' column
        grouped['a2Listing Fitment'] = grouped.apply(lambda row: row['Year'] + '|' + row['Make'] + '|' + row['Model'] + ('::' + row['Notes'] if row['Notes'] else ''), axis=1)

        # Group by 'Inventory Number' and aggregate 'a2Listing Fitment' into a list
        final_df = grouped.groupby('Inventory Number')['a2Listing Fitment'].apply(list).reset_index()

        # Convert the list of fitments into a string with "^^" as the separator
        final_df['a2Listing Fitment'] = final_df['a2Listing Fitment'].apply(lambda x: "^^".join(x))

        print(final_df)

        # Define the output file name
        output_file_name = "eBay Fitment Combine Final.xlsx"

        # Show "Complete" message when the function is done
        messagebox.showinfo("Combine", "Complete!")

        # Define the full path to the output text file on your Desktop
        output_file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
        final_df.to_excel(output_file_path, index=False, freeze_panes=(1, 0))

        # Open the Excel file and set all columns width to 15
        with xw.App(visible=False) as app:
            wb = xw.Book(output_file_path)

            # Loop through all worksheets in the workbook
            for ws in wb.sheets:
                # Loop through all columns in the worksheet
                for column in ws.api.UsedRange.Columns:
                    column.ColumnWidth = 15
                    print(f"Column {column.Column}: Width set to {column.ColumnWidth}")

            # Save the workbook if needed
            wb.save()
            print(f"Workbook saved: {output_file_path}")

            # Close the workbook
            wb.close()
            print("Workbook closed")

    except ValueError as e:
        print("Error:", e)  # Print the specific error message
        tk.messagebox.showerror("Information", f"The file you have chosen is invalid")
        return None
    except FileNotFoundError as e:
        print("Error:", e)  # Print the specific error message
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    

def Read_data_1():
    """If the file selected is valid this will load the file into the Treeview"""
    global df1  # Declare df as global to update it
    file_path = file_path_1
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df1 = pd.read_csv(excel_filename)
        else:
            df1 = pd.read_excel(excel_filename)

            # Convert the numeric columns to str
            numeric_columns = ['Year', 'Engine - CC', 'Engine - CID', 'Engine - Cylinders', 'NumDoors']
            df1[numeric_columns] = df1[numeric_columns].apply(pd.to_numeric, errors='coerce')

            print(df1.dtypes)
       
        print(df1.head(5))
        print(df1.dtypes)

        # Show "Complete" message when the function is done
        messagebox.showinfo("Read and Save", "completed successfully")
    
    # Handling error
    except ValueError as e:
        print("Error:", e)  # Print the specific error message
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"The file you have chosen is invalid")
        return None
    except FileNotFoundError as e:
        print("Error:", e)  # Print the specific error message
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"No such file as {file_path}")
        return None
    

def Read_data_2():
    """If the file selected is valid this will load the file into the Treeview"""
    global df2  # Declare df as global to update it
    file_path = file_path_2
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df2 = pd.read_csv(excel_filename)
        else:
            df2 = pd.read_excel(excel_filename)
        
            # Convert only specific columns to object
            columns_to_convert_to_object = ['Make', 'Model']
            df2[columns_to_convert_to_object] = df2[columns_to_convert_to_object].astype(object)
            
            # Convert column [Year] into int64
            df2['Year'] = df2['Year'].astype('int64')

            print(df2.dtypes)
       
        print(df2.head(5))
        print(df2.dtypes)

        # Show "Complete" message when the function is done
        messagebox.showinfo("Read and Save", "completed successfully")
    
    # Handling error
    except ValueError as e:
        print("Error:", e)  # Print the specific error message
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"The file you have chosen is invalid")
        return None
    except FileNotFoundError as e:
        print("Error:", e)  # Print the specific error message
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"No such file as {file_path}")
        return None
    

def Check_Compatibility(df2, df1):
    """Checks the compatibility of two dataframes, ignoring empty values in df2.

    Args:
        df1 (pandas.DataFrame): The first dataframe.
        df2 (pandas.DataFrame): The second dataframe.

    Returns:
        pandas.DataFrame: A DataFrame with details of incompatibility if the dataframes are not compatible, None otherwise.
    """
    
    try:
        # Check if both dataframes are loaded and not empty.
        if df2 is None or df1 is None:
            raise ValueError("Both dataframes must be provided.")

        # Find the common columns between the two dataframes.
        common_columns = set(df2.columns) & set(df1.columns)
        
        # Check if there are any common columns.
        if not common_columns:
            return pd.DataFrame([["No common columns found between the two dataframes.", None, None]], columns=['Column', 'Index', 'Value'])
        
        # Initialize a list to store messages about incompatible columns
        incompatible_columns = []

        # Check if all values in the common columns of df2 (ignoring empty values) are present in df1.
        for column in common_columns:
            values2 = df2[column].astype(str).replace('nan', '').str.strip()
            values1 = df1[column].astype(str).replace('nan', '').str.strip()

            # Filter out empty values in df2
            non_empty_values2 = values2[values2 != '']
        
            # Check if any non-empty value in df1 is not in df2.
            if not non_empty_values2.isin(values1).all():
                incompatible_values = non_empty_values2[~non_empty_values2.isin(values1)]
                for index, value in incompatible_values.items():
                    incompatible_columns.append([column, index, value])
        
        # If we found any incompatible columns, return a DataFrame with their details
        if incompatible_columns:
            return pd.DataFrame(incompatible_columns, columns=['Column', 'Index', 'Value'])
        
        # If we reach this point, all non-empty values in the common columns of df1 are present in df2, so the dataframes are compatible.
        return pd.DataFrame([["Dataframes are compatible.","", ""]], columns=['Column', 'Index', 'Value'])
    
    except ValueError as e:
        print("Error:", e)  # Print the specific error message
        messagebox.showinfo("Error", str(e))

def print_compatibility():
    compatibility = Check_Compatibility(df2, df1)  # replace df2 and df1 with your actual dataframes
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

    # Show "Complete" message when the function is done
    messagebox.showinfo("Checking", "Complete!")

    # Save the DataFrame to a CSV file
    file_path = os.path.join(os.path.expanduser("~"), "Desktop", filename)
    df.to_csv(file_path, index=False, header=True)

root.mainloop()


