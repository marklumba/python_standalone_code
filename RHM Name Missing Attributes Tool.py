from ftplib import FTP
from tkinter import messagebox
import pandas as pd
import xlwings as xw
import customtkinter
import os
from dotenv import load_dotenv

# Load the .env file
load_dotenv()

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
root.title(" RHM Name Missing Attributes Tool")

root.geometry("400x150") # set the root dimensions
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.

# Triggering buttons to call the functions
button = customtkinter.CTkButton(root, text="Download Latest File From cloudtb", command=lambda: downloadLatestFileFromFtp
                                 (localFolderPath, ftpHost, ftpUname, ftpPass, remoteFolder, substring), fg_color='blue', text_color='white', font=('Arial', 15, 'bold'))
button.grid(row=0, column=0, padx=20, pady=20)

button1 = customtkinter.CTkButton(root, text="Process Name Missing Attributes", command=lambda: process_data_from_latest_file(), 
                                  fg_color='blue', text_color='white', font=('Arial', 15, 'bold'))
button1.grid(row=1, column=0, padx=20, pady=20)


def downloadLatestFileFromFtp(localFolderPath, ftpHost, ftpUname, ftpPass, remoteWorkingDirectory, substring):
    try:
        # create an FTP client instance, use the timeout parameter for slow connections only
        ftp = FTP(timeout=60)

        # connect to the FTP server
        ftp.connect(ftpHost)

        # login to the FTP server
        ftp.login(ftpUname, ftpPass)
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

        # Derive the local file path by appending the local folder path with remote filename
        localFilePath = os.path.join(localFolderPath, latest_file)

        print("downloading file {0}".format(latest_file))

        # Download FTP file using retrbinary function
        with open(localFilePath, "wb") as file:
            ftp.retrbinary(f"RETR {latest_file}", file.write)

        # send QUIT command to the FTP server and close the connection
        ftp.quit()

        # Show "Complete" message when the function is done
        messagebox.showinfo("Download", "execution complete...!")

        return latest_file
       
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

print("execution complete...")

def process_data_from_latest_file():
    # Specify the path to your local directory where you downloaded files
    local_directory = os.path.expanduser("~/Desktop")

    # List all files in the local directory
    files_in_directory = os.listdir(local_directory)

    # Define the substring you want to filter for
    substring = "productnamebympn.xlsx"

    # Filter files to include only excel and csv files that contain the specified substring
    excel_files = [file for file in files_in_directory if (file.endswith(".xlsx") or file.endswith(".csv")) and substring in file]

    # Find the latest file based on modification time, you can use this:
    latest_file = max(excel_files, key=lambda x: os.path.getmtime(os.path.join(local_directory, x)))

    # Construct the full path to the latest file
    latest_file_path = os.path.join(local_directory, latest_file)

    try:
        if latest_file_path.endswith('.xlsx'):
            df = pd.read_excel(latest_file_path)
        else:
            df = pd.read_csv(latest_file_path)
    except EOFError:
        print(f"Failed to read file {latest_file_path}. The file might not be a valid Excel file, or it might be corrupted.")

    # # Create a mapping of AttributeValue to AttributeName
    # mapping = {}
    # for col in df.columns:
    #    if col.endswith("Value"):
    #       prefix = col.split("Value")[0]
    #       value_col = f"{prefix}Name"
    #       if value_col in df.columns:
    #          mapping[col] = df.at[0, value_col]
    
    # # Rename columns based on the mapping
    # df.rename(columns=mapping, inplace=True)
            
    # # Drop the AttributeName columns
    # df = df.loc[:, ~df.columns.str.contains('Attribute.*Name')]

    # def convert_scientific_to_standard(value):
    #    if pd.notna(value):  # Check if the value is not NaN
    #       return format(value, '.0f')
    #    else:
    #       return value  # Return the original value for NaN
       
    def convert_scientific_to_standard(value):
        if isinstance(value, (int, float)):
           return format(value, '.0f')
        else:
            return value  # or handle the error as appropriate


    # Check if 'UPC' column exists in the DataFrame
    if 'UPC' in df.columns:
       df['UPC'] = df['UPC'].apply(convert_scientific_to_standard)
       print("'UPC' column converted successfully.")
    else:
       print("'UPC' column does not exist in the DataFrame. No conversion applied.")

    # List of columns
    columns_to_drop = ['Flag', 'FlagDescription', 'Blocked Comment']

    # Check if columns exist in DataFrame before dropping
    columns_exist = all(col in df.columns for col in columns_to_drop)

    if columns_exist:
       df.drop(columns_to_drop, axis=1, inplace=True)
       print("Columns dropped successfully.")
    else:
       print("Columns do not exist in the DataFrame. No columns were dropped.")

    # Sort by 'Product-Name' and drop rows with missing 'Product-Name'
    df = df.sort_values(by='Product-Name')
    df = df[df['Product-Name'].notna()]

    # Identify columns with missing values for each 'Product-Name'
    grouped = df.groupby('Product-Name')
    
    # create an empty dictionary
    missing_data = {}

    # Iterate on the grouped
    for name, group in grouped:
        # Find columns with any missing values
        missing_cols = group.columns[group.isnull().any()].tolist()
    
        # Ignore columns where all values are missing
        missing_cols = [col for col in missing_cols if group[col].notna().any()]
    
        # Add to the dictionary
        missing_data[name] = missing_cols

    # Convert the dictionary to a DataFrame
    missing_data_df = pd.DataFrame(list(missing_data.items()), columns=['Product-Name', 'Columns with Missing Data'])

    # Remove '[]' at the values in column 'Columns with Missing Data'
    missing_data_df['Columns with Missing Data'] =  missing_data_df['Columns with Missing Data'].apply(lambda x: str(x).replace('[','').replace(']',''))

    # Merge the original DataFrame with the missing_data_df DataFrame
    merged_df = pd.merge(df, missing_data_df, on='Product-Name', how='left')

    # Export the transformed data to a new Excel file
    file_path = os.path.join(os.path.expanduser("~"), "Desktop", "Name Missing Attributes.xlsx")
    # Write DataFrame to Excel
    merged_df.to_excel(file_path, index=False, freeze_panes=(1, 2))

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
    
    # Show "Complete" message when the function is done
    messagebox.showinfo("Process", "Complete!")

root.mainloop()











