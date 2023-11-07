import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import xlwings as xw



# initalise the tkinter GUI
root = tk.Tk()
root.title("eBay_MVL_Creator - Make Filter Tool")


root.geometry("600x600") # set the root dimensions
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.

# Declare df as a global variable
df = None

# Frame for TreeView
frame1 = tk.LabelFrame(root, text="Excel Data")
frame1.place(height=300, width=600)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File")
file_frame.place(height=200, width=600, rely=0.65, relx=0)

# Buttons
button1 = tk.Button(file_frame, text="Browse A File", command=lambda: File_dialog())
button1.place(rely=0.45, relx=0.25)

button2 = tk.Button(file_frame, text="MVL_1970", command=lambda: Transform_1())
button2.place(rely=0.65, relx=0.70)

button3 = tk.Button(file_frame, text="MVL_1990", command=lambda: Transform_2())
button3.place(rely=0.65, relx=0.55)

button4 = tk.Button(file_frame, text="MVL_2000", command=lambda: Transform_3())
button4.place(rely=0.65, relx=0.40)

button5 = tk.Button(file_frame, text="MVL_2010", command=lambda: Transform_4())
button5.place(rely=0.65, relx=0.25)

button6 = tk.Button(file_frame, text="MVL_Filter_1970", command=lambda: Transform_5())
button6.place(rely=0.85, relx=0.25)

button7 = tk.Button(file_frame, text="MVL_Filter_1990", command=lambda: Transform_6())
button7.place(rely=0.85, relx=0.45)

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)

## Treeview Widget
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

            # Delete column [Region]
            #df = df.drop(columns=["Region"])

            if "Region" in df.columns:
                df = df.drop(columns=["Region"])
            else:
                print("Column 'Region' does not exist in df.")

            # Delete all years and rows of data before i.e. 1970.
            df = df[df["Year"].between(1970, 2024)]


            # Define a list of valid makes
            valid_makes = [
            'Acura', 'Alfa Romeo', 'Aston Martin', 'Audi', 'Bentley', 'BMW', 'Buick',
            'Cadillac', 'Chevrolet', 'Chrysler', 'Daihatsu', 'Dodge', 'Eagle', 'Ferrari',
            'Fiat', 'Ford', 'Freightliner', 'Genesis', 'Geo', 'GMC', 'Hino', 'Honda',
            'Hummer', 'Hyundai', 'INFINITI', 'Isuzu', 'Jaguar', 'Jeep', 'Kenworth', 'Kia',
            'Lamborghini', 'Land Rover', 'Lexus', 'Lincoln', 'Lotus', 'Mack', 'Maserati',
            'Mazda', 'Mercedes-Benz', 'Mercury', 'Mini', 'Mitsubishi', 'Mobility Ventures',
            'Nissan', 'Oldsmobile', 'Peterbilt', 'Plymouth', 'Pontiac', 'Porsche', 'Ram',
            'Saab', 'Saturn', 'Scion', 'Smart', 'Sterling Truck', 'Subaru', 'Suzuki',
            'Tesla', 'Toyota', 'Vespa', 'Volkswagen', 'Volvo', 'VPG', 'Western Star', 'Rivian',
            'Lucid', 'Polestar'
             ]

            # Filter "Make" column based on the valid makes list
            df = df[df['Make'].isin(valid_makes)]
            
            # Export the transformed data to a new Excel file
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", "eBay_Master_Vehicle_List_1970.xlsx")
            df.to_excel(file_path, index=False,freeze_panes=(1, 1))

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
    
def Transform_2():
    """If the file selected is valid this will load the file into the Treeview"""
    global df  # Declare df as global to update it
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)
            
            # Delete all years and rows of data before i.e. 1990.
            df = df[df["Year"].between(1990, 2024)]

            # Define a list of valid makes
            valid_makes = [
            'Acura', 'Alfa Romeo', 'Aston Martin', 'Audi', 'Bentley', 'BMW', 'Buick',
            'Cadillac', 'Chevrolet', 'Chrysler', 'Daihatsu', 'Dodge', 'Eagle', 'Ferrari',
            'Fiat', 'Ford', 'Freightliner', 'Genesis', 'Geo', 'GMC', 'Hino', 'Honda',
            'Hummer', 'Hyundai', 'INFINITI', 'Isuzu', 'Jaguar', 'Jeep', 'Kenworth', 'Kia',
            'Lamborghini', 'Land Rover', 'Lexus', 'Lincoln', 'Lotus', 'Mack', 'Maserati',
            'Mazda', 'Mercedes-Benz', 'Mercury', 'Mini', 'Mitsubishi', 'Mobility Ventures',
            'Nissan', 'Oldsmobile', 'Peterbilt', 'Plymouth', 'Pontiac', 'Porsche', 'Ram',
            'Saab', 'Saturn', 'Scion', 'Smart', 'Sterling Truck', 'Subaru', 'Suzuki',
            'Tesla', 'Toyota', 'Vespa', 'Volkswagen', 'Volvo', 'VPG', 'Western Star', 'Rivian',
            'Lucid', 'Polestar'
             ]

            # Filter "Make" column based on the valid makes list
            df = df[df['Make'].isin(valid_makes)]

            # Export the transformed data to a new Excel file
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", "eBay_Master_Vehicle_List_1990.xlsx")
            df.to_excel(file_path, index=False, freeze_panes=(1, 1))

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
    
def Transform_3():
    """If the file selected is valid this will load the file into the Treeview"""
    global df  # Declare df as global to update it
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

            # Delete all years and rows of data before i.e. 2000.
            df = df[df["Year"].between(2000, 2024)]

            # Define a list of valid makes
            valid_makes = [
            'Acura', 'Alfa Romeo', 'Aston Martin', 'Audi', 'Bentley', 'BMW', 'Buick',
            'Cadillac', 'Chevrolet', 'Chrysler', 'Daihatsu', 'Dodge', 'Eagle', 'Ferrari',
            'Fiat', 'Ford', 'Freightliner', 'Genesis', 'Geo', 'GMC', 'Hino', 'Honda',
            'Hummer', 'Hyundai', 'INFINITI', 'Isuzu', 'Jaguar', 'Jeep', 'Kenworth', 'Kia',
            'Lamborghini', 'Land Rover', 'Lexus', 'Lincoln', 'Lotus', 'Mack', 'Maserati',
            'Mazda', 'Mercedes-Benz', 'Mercury', 'Mini', 'Mitsubishi', 'Mobility Ventures',
            'Nissan', 'Oldsmobile', 'Peterbilt', 'Plymouth', 'Pontiac', 'Porsche', 'Ram',
            'Saab', 'Saturn', 'Scion', 'Smart', 'Sterling Truck', 'Subaru', 'Suzuki',
            'Tesla', 'Toyota', 'Vespa', 'Volkswagen', 'Volvo', 'VPG', 'Western Star', 'Rivian',
            'Lucid', 'Polestar'
             ]
            
            # Filter "Make" column based on the valid makes list
            df = df[df['Make'].isin(valid_makes)]


            # Export the transformed data to a new Excel file
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", "eBay_Master_Vehicle_List_2000.xlsx")
            df.to_excel(file_path, index=False, freeze_panes=(1, 1))

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

def Transform_4():
    """If the file selected is valid this will load the file into the Treeview"""
    global df  # Declare df as global to update it
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

            # Delete all years and rows of data before i.e. 2010.
            df = df[df["Year"].between(2010, 2024)]

            # Define a list of valid makes
            valid_makes = [
            'Acura', 'Alfa Romeo', 'Aston Martin', 'Audi', 'Bentley', 'BMW', 'Buick',
            'Cadillac', 'Chevrolet', 'Chrysler', 'Daihatsu', 'Dodge', 'Eagle', 'Ferrari',
            'Fiat', 'Ford', 'Freightliner', 'Genesis', 'Geo', 'GMC', 'Hino', 'Honda',
            'Hummer', 'Hyundai', 'INFINITI', 'Isuzu', 'Jaguar', 'Jeep', 'Kenworth', 'Kia',
            'Lamborghini', 'Land Rover', 'Lexus', 'Lincoln', 'Lotus', 'Mack', 'Maserati',
            'Mazda', 'Mercedes-Benz', 'Mercury', 'Mini', 'Mitsubishi', 'Mobility Ventures',
            'Nissan', 'Oldsmobile', 'Peterbilt', 'Plymouth', 'Pontiac', 'Porsche', 'Ram',
            'Saab', 'Saturn', 'Scion', 'Smart', 'Sterling Truck', 'Subaru', 'Suzuki',
            'Tesla', 'Toyota', 'Vespa', 'Volkswagen', 'Volvo', 'VPG', 'Western Star', 'Rivian',
            'Lucid', 'Polestar'
             ]

            # Filter "Make" column based on the valid makes list
            df = df[df['Make'].isin(valid_makes)]

            # Export the transformed data to a new Excel file
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", "eBay_Master_Vehicle_List_2010.xlsx")
            df.to_excel(file_path, index=False, freeze_panes=(1, 1))

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
    

def Transform_5():
    """If the file selected is valid this will load the file into the Treeview"""
    global df  # Declare df as global to update it
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

            # Delete column [Region] and [ePID]
            #df = df.drop(columns=["ePID"])

            if "ePID" in df.columns:
               df = df.drop(columns=["ePID"])
            else:
               print("Column 'ePID' does not exist in df.")

            # Delete all years and rows of data before i.e. 1970.
            df = df[df["Year"].between(1970, 2024)]

            # Define a list of valid makes
            valid_makes = [
            'Acura', 'Alfa Romeo', 'Aston Martin', 'Audi', 'Bentley', 'BMW', 'Buick',
            'Cadillac', 'Chevrolet', 'Chrysler', 'Daihatsu', 'Dodge', 'Eagle', 'Ferrari',
            'Fiat', 'Ford', 'Freightliner', 'Genesis', 'Geo', 'GMC', 'Hino', 'Honda',
            'Hummer', 'Hyundai', 'INFINITI', 'Isuzu', 'Jaguar', 'Jeep', 'Kenworth', 'Kia',
            'Lamborghini', 'Land Rover', 'Lexus', 'Lincoln', 'Lotus', 'Mack', 'Maserati',
            'Mazda', 'Mercedes-Benz', 'Mercury', 'Mini', 'Mitsubishi', 'Mobility Ventures',
            'Nissan', 'Oldsmobile', 'Peterbilt', 'Plymouth', 'Pontiac', 'Porsche', 'Ram',
            'Saab', 'Saturn', 'Scion', 'Smart', 'Sterling Truck', 'Subaru', 'Suzuki',
            'Tesla', 'Toyota', 'Vespa', 'Volkswagen', 'Volvo', 'VPG', 'Western Star', 'Rivian',
            'Lucid', 'Polestar'
             ]

            # Filter "Make" column based on the valid makes list
            df = df[df['Make'].isin(valid_makes)]
            
            # Export the transformed data to a new Excel file
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", "eBay_MVL_Filter_1970.xlsx")
            df = df[["Year","Make","Model","Submodel","Body","Fuel Type Name","NumDoors","Drive Type","Engine - Liter_Display","Engine - Cylinders","Trim","Parts Model","KBB_MODEL","Aspiration","Cylinder Type Name","DisplayName","Engine","Engine - Block Type","Engine - CC","Engine - CID"]]
            df.to_excel(file_path, index=False, freeze_panes=(1, 1))

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

def Transform_6():
    """If the file selected is valid this will load the file into the Treeview"""
    global df  # Declare df as global to update it
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

            # Delete all years and rows of data before i.e. 1990.
            df = df[df["Year"].between(1990, 2024)]

            # Define a list of valid makes
            valid_makes = [
            'Acura', 'Alfa Romeo', 'Aston Martin', 'Audi', 'Bentley', 'BMW', 'Buick',
            'Cadillac', 'Chevrolet', 'Chrysler', 'Daihatsu', 'Dodge', 'Eagle', 'Ferrari',
            'Fiat', 'Ford', 'Freightliner', 'Genesis', 'Geo', 'GMC', 'Hino', 'Honda',
            'Hummer', 'Hyundai', 'INFINITI', 'Isuzu', 'Jaguar', 'Jeep', 'Kenworth', 'Kia',
            'Lamborghini', 'Land Rover', 'Lexus', 'Lincoln', 'Lotus', 'Mack', 'Maserati',
            'Mazda', 'Mercedes-Benz', 'Mercury', 'Mini', 'Mitsubishi', 'Mobility Ventures',
            'Nissan', 'Oldsmobile', 'Peterbilt', 'Plymouth', 'Pontiac', 'Porsche', 'Ram',
            'Saab', 'Saturn', 'Scion', 'Smart', 'Sterling Truck', 'Subaru', 'Suzuki',
            'Tesla', 'Toyota', 'Vespa', 'Volkswagen', 'Volvo', 'VPG', 'Western Star', 'Rivian',
            'Lucid', 'Polestar'
             ]

            # Filter "Make" column based on the valid makes list
            df = df[df['Make'].isin(valid_makes)]
            
            # Export the transformed data to a new Excel file
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", "eBay_MVL_Filter_1990.xlsx")
            df = df[["Year","Make","Model","Submodel","Body","Fuel Type Name","NumDoors","Drive Type","Engine - Liter_Display","Engine - Cylinders","Trim","Parts Model","KBB_MODEL","Aspiration","Cylinder Type Name","DisplayName","Engine","Engine - Block Type","Engine - CC","Engine - CID"]]
            df.to_excel(file_path, index=False, freeze_panes=(1, 1))

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
