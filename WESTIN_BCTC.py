from ftplib import FTP
import pandas as pd
import xlwings as xw
import os
from dotenv import load_dotenv
from ftplib import FTP_TLS
from datetime import datetime
import datetime


# Load the .env file
load_dotenv()

# Usage credentials
localFolderPath = os.path.expanduser("~/Desktop") # Specify folder path save
remoteFolder = '/OSA/SDC/WESTIN_BCTC/ACES' # Specify folder path
remoteFolder_2 = '/OSA/Vendor_Feeds/MEY' # Specify folder path
substring = 'SDC_BCTC_ACES'  # Specify the substring to filter files
substring_2 = 'Meyer Inventory.csv'  # Specify the substring to filter files
substring_3 = 'SDC_BCTC_ACES Flat Export.xlsx' # Specify the substring to filter files

# Get the variables
ftpUname = os.getenv('ftpUname')
ftpPass = os.getenv('ftpPass')
ftpHost = os.getenv('ftpHost')

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

        # Read the Excel file into a pandas DataFrame
        try:
           # Read the CSV file in chunks
            reader = pd.read_csv(localFilePath, sep="|", low_memory=False) 

            # Iterate over the columns
            for column in reader.columns: 
                # If all values in the column are NaN (empty), drop the column
                if reader[column].isna().all():
                   reader.drop(column, axis=1, inplace=True)   
               

            # Export the transformed data to a new Excel file
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", "SDC_BCTC_ACES_OUTPUT.xlsx")
            reader.to_excel(file_path, index=False, freeze_panes=(1, 0))
      
        
            print(f'Saved DataFrame to {file_path}')  # prints the file path
           

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


            print(reader.head())  # prints the first 5 rows of each chunk
        except Exception as e:
            print(f"An error occurred: {str(e)}")
       
        # send QUIT command to the FTP server and close the connection
        ftp.quit()
       
        return latest_file
       
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None
    
def downloadLatestFileFromFtp_2(localFolderPath, ftpHost, ftpUname, ftpPass, remoteWorkingDirectory, substring_2):
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
        matching_files = [file for file in file_list if substring_2 in file]

        if not matching_files:
            print(f"No files found containing the substring: {substring_2}")
            return None 

        # Find the latest file (based on modification time) and (substrings)
        latest_file_2 = max(matching_files, key=lambda x: ftp.sendcmd(f"MDTM {x}").split()[1])
      
        # Derive the local file path by appending the local folder path with remote filename
        localFilePath = os.path.join(localFolderPath, latest_file_2)

        print("downloading file {0}".format(latest_file_2))

        # Download FTP file using retrbinary function
        with open(localFilePath, "wb") as file:
            ftp.retrbinary(f"RETR {latest_file_2}", file.write)

        # Read the Excel file into a pandas DataFrame
        try:
            columns_to_load = ['MFGName', 'MFG Item Number', 'Available']
            # Read the CSV file in chunks
            df1 = pd.read_csv(localFilePath, usecols=columns_to_load) 

            # Filter the DataFrame
            df1 = df1[df1['MFGName'] == 'Westin Automotive']   
               

            # Export the transformed data to a new Excel file
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", "Meyer Inventory.xlsx")
            df1.to_excel(file_path, index=False, freeze_panes=(1, 0))
      
        
            print(f'Saved DataFrame to {file_path}')  # prints the file path
           

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


            print(df1.head())  # prints the first 5 rows of each chunk
        except Exception as e:
            print(f"An error occurred: {str(e)}")
       
        # send QUIT command to the FTP server and close the connection
        ftp.quit()
       
        return latest_file_2
       
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

    
# Call the function with the provided parameters
latest_file = downloadLatestFileFromFtp(localFolderPath, ftpHost, ftpUname, ftpPass, remoteFolder, substring)
latest_file_2 = downloadLatestFileFromFtp_2(localFolderPath, ftpHost, ftpUname, ftpPass, remoteFolder_2, substring_2)


# Paths to the Excel files
file_path1 = os.path.join(os.path.expanduser("~"), "Desktop", "SDC_BCTC_ACES_OUTPUT.xlsx")
file_path2 = os.path.join(os.path.expanduser("~"), "Desktop", "Meyer Inventory.xlsx")

# Read the Excel files into pandas DataFrames
df1 = pd.read_excel(file_path1,  engine='openpyxl')
df2 = pd.read_excel(file_path2,  engine='openpyxl')

# Merge the DataFrames
merged_df = pd.merge(df1, df2, how='left', left_on='Part', right_on='MFG Item Number')

# Rename the column 'Available' into 'Inventory'
merged_df.rename(columns={'Available': 'Inventory'}, inplace=True)

# Drop columns 'MFGName', 'MFG Item Number'
merged_df.drop(columns=['MFGName', 'MFG Item Number'], inplace=True)

# Columns list for new order
new_order = ['AAIA_BrandID', 'Part', 'Inventory', 'Year', 'Make', 'Model', 'Submodel', 'PartType', 'Position', 'Quantity', 'Region',
             'FitmentNotes', 'MfrLabel', 'Liter', 'Cylinders', 'BlockType', 'FuelTypeName', 'CC', 'CID', 'EngineBoreInches', 'EngineBoreMetric',
             'EngineStrokeInches', 'EngineStrokeMetric', 'BodyTypeName', 'BodyNumberOfDoors', 'BedTypeName', 'BedLengthInches', 'BedLengthMetric',
             'DriveTypeName']

merged_df = merged_df.reindex(columns=new_order)

# Sort by 'Part' ascending
merged_df = merged_df.sort_values(by='Part', ascending=True)


# Define the output file name with the date and time
output_file_name = f"SDC_BCTC_ACES Flat Export.xlsx"

# Export the transformed data to a new Excel file
file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
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

# Now merged_df contains the result of the merge operation
print(merged_df.head())


def upload_file(localFilePath, ftpHost, ftpUname, ftpPass, remoteDirectory):
    try:
        # Create an FTP client instance
        with FTP_TLS() as ftp:
            ftp.connect(ftpHost)
            ftp.login(ftpUname, ftpPass)
            ftp.set_pasv(True)  # Enable passive mode
            print("Successfully connected and logged in to the FTP server.")

            # Change working directory if specified
            if remoteDirectory:
                ftp.cwd(remoteDirectory)

            # Extract the filename from the localFilePath
            fileName = os.path.basename(localFilePath)

            # Format the date as a string in the format 'YYYY-MM-DD'
            date_string = datetime.datetime.now().strftime("%Y-%m-%d")

            # Construct the new file name
            new_file_name = f"{fileName[:-5]}_{date_string}.xlsx"  # assuming the filename ends with '.xlsx'
            remoteFilePath = new_file_name

            print(f"Uploading file as {new_file_name}")

            # Upload file to FTP server
            with open(localFilePath, 'rb') as file:
                ftp.storbinary(f'STOR {remoteFilePath}', file)

            print("File uploaded and renamed successfully.")

            # Close the FTP connection
            ftp.quit()

    except FileNotFoundError:
        print("Local file not found. Please check the file path.")
    except Exception as e:
        print(f"An FTP error occurred: {str(e)}")



# Define the local file path
localFilePath = os.path.join(localFolderPath, substring_3)

# Call the function to upload the file directly to 'ACES' directory
upload_file(localFilePath, ftpHost, ftpUname, ftpPass, remoteFolder)









