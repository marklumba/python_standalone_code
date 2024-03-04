from ftplib import FTP_TLS
from tkinter import messagebox
import pandas as pd
import xlwings as xw
import customtkinter
import tkinter as tk
from tkinter import filedialog, ttk
import os
from dotenv import load_dotenv
import pathlib
from pathlib import Path




# Load the .env file
load_dotenv()


# Declare global variables for file paths
file_path_1 = None
file_path_2 = None

# Declare df1 and df2 as global variables
df1 = None
df2 = None

# Setting up theme of the app
customtkinter.set_appearance_mode("system")

# Setting up them of your components
customtkinter.set_default_color_theme("blue")

# Usage credentials
localFolderPath = os.path.expanduser("~/Desktop") # Specify folder path save
remoteFolder = '/OSA/ChannelAdvisor/AutomatedExports' # Specify folder path
substring = 'productnamebympn.xlsx'  # Specify the substring to filter files

# Get the variables
ftpUname = os.getenv('ftpUname')
ftpPass = os.getenv('ftpPass')
ftpHost = os.getenv('ftpHost')

# initalise the tkinter GUI
root = customtkinter.CTk()
root.title("RHM Inventory Copy Tool")

root.geometry("800x700") # set the root dimensions
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.

frame1 = tk.LabelFrame(root, text="Functions", bg="lightgrey", fg="black", font=("Arial", 12, "bold"))
frame1.place(height=150, width=800, relx=0.01)

frame2 = tk.LabelFrame(root, text="Download and Read File", bg="lightgrey", fg="black", font=("Arial", 12, "bold"))
frame2.place(height=200, width=800, rely=0.30, relx=0.01)

frame3 = tk.LabelFrame(root, text="Display File", bg="lightgrey", fg="black", font=("Arial", 12, "bold"))
frame3.place(height=200, width=800, rely=0.65, relx=0.01)

button1 = customtkinter.CTkButton(frame2, text="Download Latest File From cloudtb", command=lambda: downloadLatestFileFromFtp
                                  (ftpHost, ftpUname, ftpPass, remoteFolder, substring), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button1.place(rely=0.2, relx=0.01)

button2 = customtkinter.CTkButton(frame3, text="Select CA Inventory Copy Template", command=lambda: file_dialog_1
                                  (), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button2.place(rely=0.2, relx=0.01)

button3 = customtkinter.CTkButton(frame2, text="Read File and Save File Path Link", command=lambda: read_data_1
                                  (), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button3.place(rely=0.4, relx=0.01)

button4 = customtkinter.CTkButton(frame3, text="Read File and Save File Path Link", command=lambda: read_data_2
                                  (), fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button4.place(rely=0.4, relx=0.01)

button5 = customtkinter.CTkButton(frame1, text="CA Inventory Copy", command=lambda: ca_inventory_copy_tool
                                  (df1, df2), 
                                  fg_color='blue', text_color='white', font=('Arial', 12, 'bold'))
button5.place(rely=0.2, relx=0.01)


label_file1 = ttk.Label(frame3, text="No File Selected", background="lightgrey", foreground="blue", font=("Arial", 11, "bold"))
label_file1.place(rely=0, relx=0)



def file_dialog_1():
    """This Function will open the file explorer and assign the chosen file path to label_file2"""
    global file_path_2  # Declare as global to update it
    try:
        filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
        label_file1["text"] = filename
        file_path_2 = filename  # Update the global variable
    except Exception as e:
        print("An error occurred:", str(e))
    return filename


def read_data_1():
    global df1  # Declare df1 as global to update it
    
    # Specify the path to your local directory where you downloaded files
    local_directory = pathlib.Path.home() / 'Desktop'

    # List all files in the local directory
    files_in_directory = os.listdir(local_directory)

    # Define the substring you want to filter for
    substring = "productnamebympn.xlsx"
  
    # Filter files to include only Excel and CSV files that contain the specified substring
    relevant_files = [file for file in files_in_directory if file.endswith((".xlsx", ".csv")) and substring in file]

    if not relevant_files:
        # No matching files found
        print(f"No files found containing the substring: {substring}")
        return None

    # Find the latest file based on modification time
    latest_file = max(relevant_files, key=lambda x: (pathlib.Path(local_directory) / x).stat().st_mtime)

    # Construct the full path to the latest file
    latest_file_path = pathlib.Path(local_directory) / latest_file

    try:
        # Read the latest file based on its extension
        if latest_file_path.suffix == '.xlsx':
            df1 = pd.read_excel(latest_file_path)
        else:
            df1 = pd.read_csv(latest_file_path, encoding='latin-1')  # change 'latin-1' to the correct encoding
    except pd.errors.EmptyDataError as e:
        print(f"File {latest_file_path} is empty. Error: {str(e)}")
        return None
    except Exception as e:
        print(f"An error occurred while reading file {latest_file_path}. Error: {str(e)}")
        return None

    print(df1.head(5))

    # Show "Complete" message when the function is done
    messagebox.showinfo("Read and Save", "completed successfully!")
       


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
            
                    
        print(df2.head(5))

        # Show "Complete" message when the function is done
        messagebox.showinfo("Read and Save", "completed successfully!")
           
    except ValueError as e:
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"The file you have chosen is invalid: {str(e)}")
        return None
    except FileNotFoundError as e:
        # Display an error message using the messagebox
        messagebox.showerror("Error", f"No such file as {file_path}: {str(e)}")
        return None
    
  
def ca_inventory_copy_tool(df1, df2):

    # Merge df1 and df2 on the common columns
    df_new = pd.merge(df2, df1, how='left', left_on='Copy Inventory Number', right_on='Inventory Number')
     
    # Drop column 'MPN_y' and Inventory Number
    df_new.drop(['MPN_y', 'Inventory Number'], axis=1, inplace=True)

    # Rename columns using a dictionary
    df_new.rename(columns={'MPN_x': 'MPN', 'New Inventory Number': 'Inventory Number'}, inplace=True)

    # Assign default values to new columns using a dictionary
    df_new = df_new.assign(Flag='YellowFlag', FlagDescription='Copy Tool - Needs Work/Pre-Launch', Blocked='TRUE', **{'Website-Store-Mapping': 'DNI'})


    def convert_scientific_to_standard(value):
        try:
           # Convert the input value to a float
           value = float(value)
           return format(value, '.0f')
        except ValueError as e:
           print(f"Error converting {value} to float: {e}")
           return "Not a valid number"


    # Check if 'UPC' column exists in the DataFrame
    if 'UPC' in df_new.columns:
       df_new['UPC'] = df_new['UPC'].apply(convert_scientific_to_standard)
       print("'UPC' column converted successfully.")
    else:
       print("'UPC' column does not exist in the DataFrame. No conversion applied.")
       
    # Export the filtered data to a new Excel file
    file_path = os.path.join(os.path.expanduser("~"), "Desktop", "CA Inventory Copy Output.xlsx")
    
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter') 

    # Write each dataframe to a different worksheet.
    df_new.to_excel(writer, sheet_name='CA Inventory Copy Ouput', index=False, freeze_panes=(1, 1))

    # Close the Pandas Excel writer and output the Excel file.
    writer.close()
    
    # Open the Excel file and set all columns width to 15
    with xw.App(visible=False) as app:
       wb = xw.Book(file_path)

       # Loop through all worksheets in the workbook
       for ws in wb.sheets:
           print(f"Processing sheet: {ws.name}")  # Print the name of the sheet being processed

           try:
              # Loop through all columns in the worksheet
               for column in ws.api.UsedRange.Columns:
                    column.ColumnWidth = 15
           except Exception as e:
                print(f"An error occurred while processing sheet {ws.name}: {str(e)}")  # Print any errors that occur

       # Save the workbook if needed
       wb.save()
       
       # Close the workbook
       wb.close()

    # Show "Complete" message when the function is done
    messagebox.showinfo("Process", "complete")
       

def downloadLatestFileFromFtp(ftpHost, ftpUname, ftpPass, remoteWorkingDirectory, substring):
    try:
        # create an FTP_TLS client instance, use the timeout parameter for slow connections only
        ftp = FTP_TLS(timeout=60)
        
        # connect to the FTP server
        ftp.connect(ftpHost)
        ftp.login(ftpUname, ftpPass)
        
        # Switch to secure data connection
        ftp.prot_p()

        print("Successfully connected and logged in to the FTP server.")

        # change current working directory if specified
        if not (remoteWorkingDirectory == None or remoteWorkingDirectory.strip() == ""):
            _ = ftp.cwd(remoteWorkingDirectory)

        # List files in the remote directory
        file_list = ftp.nlst()

        # Filter files based on the specified substring
        matching_files = [file for file in file_list if substring in file]

        if not matching_files:
            print(f"No files found containing the substring: {substring}")
            return None 

        # Find the latest file (based on modification time) and (substrings)
        latest_file = max(matching_files, key=lambda x: ftp.sendcmd(f"MDTM {x}").split()[1])
      
        print("downloading file {0}".format(latest_file))

        # Download FTP file using retrbinary function
        local_file_path = os.path.join(Path.home(), 'Desktop', latest_file)
        with open(local_file_path, 'wb') as local_file:
            ftp.retrbinary(f"RETR {latest_file}", local_file.write)

        # send QUIT command to the FTP server and close the connection
        ftp.quit()

        # Show "Complete" message when the function is done
        print("File downloaded successfully!")
        # Show "Complete" message when the function is done
        messagebox.showinfo("Process", "File downloaded successfully!")
        return local_file_path
       
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None


print("execution complete...")

root.mainloop()

