from ftplib import FTP
import pandas as pd
import os
import xlwings as xw

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

        return latest_file
    
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

print("execution complete...")

# Usage credentials
ftpHost = ''
ftpUname = ''
ftpPass = ''
localFolderPath = os.path.expanduser("~/Desktop")
remoteFolder = '/OSA/ChannelAdvisor/AutomatedExports'
substring = 'productnamebympn.xlsx'  # Specify the substring to filter files

# Call the function with the provided parameters
latest_file = downloadLatestFileFromFtp(localFolderPath, ftpHost, ftpUname, ftpPass, remoteFolder, substring)

print(f"The latest file downloaded is: {latest_file}")

# Specify the path to your local directory where you downloaded files
local_directory = os.path.expanduser("~/Desktop")

# List all files in the local directory
files_in_directory = os.listdir(local_directory)

# Define the substring you want to filter for
substring = "productnamebympn.xlsx"

# Filter files to include only excel and csv files that contain the specified substring
excel_files = [file for file in files_in_directory if (file.endswith(".xlsx") or file.endswith(".csv")) and substring in file]

# If you want to find the latest file based on modification time, you can use this:
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

# Create a mapping of AttributeValue to AttributeName
mapping = {}
for col in df.columns:
    if col.endswith("Value"):
       prefix = col.split("Value")[0]
       value_col = f"{prefix}Name"
       if value_col in df.columns:
           mapping[col] = df.at[0, value_col]


# Rename columns based on the mapping
df.rename(columns=mapping, inplace=True)
            
# Drop the AttributeName columns
df = df.loc[:, ~df.columns.str.contains('Attribute.*Name')]

# Define a custom function to convert scientific notation to standard notation
def convert_scientific_to_standard(value):
    if pd.notna(value):  # Check if the value is not NaN
        return format(value, '.0f')
    else:
        return value  # Return the original value for NaN

# Apply the function to the DataFrame
df['UPC'] = df['UPC'].apply(convert_scientific_to_standard)

# Drop columns Flag, FlagDescription and Blocked Comment
df.drop(['Flag', 'FlagDescription', 'Blocked Comment'], axis=1, inplace=True)

# Sort by 'Product-Name' and drop rows with missing 'Product-Name'
df = df.sort_values(by='Product-Name')
df = df[df['Product-Name'].notna()]

# Identify columns with missing values for each 'Product-Name'
grouped = df.groupby('Product-Name')

missing_data = {}
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

print("Script has completed successfully.")










