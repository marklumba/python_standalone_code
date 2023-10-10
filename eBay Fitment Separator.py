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

# Declare global variables for file paths
file_path_1 = None

# Setting up theme of the app
ctk.set_appearance_mode("system")

# Setting up them of your components
ctk.set_default_color_theme("blue")

# Initialize the tkinter GUI
root = ctk.CTk()
root.title("eBay Fitment Separator")

root.geometry("600x400")  # set the root dimensions
root.pack_propagate(False)  # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0)  # makes the root window fixed in size.

# Frame for TreeView
frame1 = tk.LabelFrame(root, text="Excel Data", bg="lightgrey", fg="black", font=("Arial", 15, "bold"))
frame1.place(height=200, width=750)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File", bg="lightgrey", fg="black", font=("Arial", 15, "bold"))
file_frame.place(height=150, width=750, rely=0.65, relx=0)

# Buttons
button1 = ctk.CTkButton(file_frame, text="Browse for File", command=lambda: File_dialog(), fg_color='blue',
                                  text_color='white', font=('Arial', 15, 'bold'))
button1.place(rely=0.2, relx=0.01)

button2 = ctk.CTkButton(file_frame, text="Run Transformation", command=lambda: Transform_1(), fg_color='blue',
                                  text_color='white', font=('Arial', 15, 'bold'))
button2.place(rely=0.6, relx=0.01)

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected", background="lightgrey", foreground="blue",
                       font=("Arial", 12, "bold"))
label_file.place(rely=0, relx=0)


# Treeview Widget
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1)  # set the height and width of the widget to 100% of its container (frame1).

treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview)  # command means update the yaxis view of the widget
treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview)  # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)  # assign the scrollbars to the Treeview Widget
treescrollx.pack(side="bottom", fill="x")  # make the scrollbar fill the x axis of the Treeview widget
treescrolly.pack(side="right", fill="y")  # make the scrollbar fill the y axis of the Treeview widget


def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file1"""
    global file_path_1  # Declare as global to update it
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file["text"] = filename
    file_path_1 = filename  # Update the global variable
    return None

def Transform_1():
    """If the file selected is valid, this will load the file into the Treeview"""
    global df  # Declare df as global to update it
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename, nrows=5)
        else:
            df = pd.read_excel(excel_filename, nrows=5)

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
            fitments = fitment.split("^^")

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
                make_model, notes = make_model_notes.split("::") if "::" in make_model_notes else (make_model_notes, '')
                make, model = make_model.split("|") if "|" in make_model else (make_model, '')

                # Iterate through each year in the year range
                for year in range(start_year, end_year + 1):
                    # Append the transformed line to the list of transformed data
                    transformed_data.append([inventory_number, str(year), make, model, notes])

        # Convert list of lists to DataFrame
        final_df = pd.DataFrame(transformed_data, columns=['Inventory Number', 'Year', 'Make', 'Model', 'Notes'])

        # Define the output file name
        output_file_name = "eBayFitmentSeparatorFinal.xlsx"

        # Show "Complete" message when the function is done
        messagebox.showinfo("eBay_Fitment_Separator", "Complete!")

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
        tv1.heading(column, text=column)  # let the column heading = column name

    df_rows = df.to_numpy().tolist()  # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end", values=row)  # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
    return None

def clear_data():
    tv1.delete(*tv1.get_children())
    return None

root.mainloop()



