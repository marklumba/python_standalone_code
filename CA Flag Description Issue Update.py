import ftplib
import os
import pandas as pd
from datetime import datetime
#from ftplib import FTP


def downloadLatestFileFromFtp(localFolderPath, ftpHost, ftpUname, ftpPass, remoteWorkingDirectory, substring):
    # create an FTP client instance, use the timeout parameter for slow connections only
    ftp = ftplib.FTP(timeout=60)

    # connect to the FTP server
    ftp.connect(ftpHost)

    # login to the FTP server
    ftp.login(ftpUname, ftpPass)

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

# Usage example credentials
ftpHost = ''
ftpUname = ''
ftpPass = ''
localFolderPath = os.path.expanduser("~/Desktop")
remoteFolder = "/Reports"
substring = "Flag Description"  # Specify the substring to filter files

latest_file = downloadLatestFileFromFtp(localFolderPath, ftpHost, ftpUname, ftpPass, remoteFolder, substring)

print(f"The latest file downloaded is: {latest_file}")

# Specify the path to your local directory where you downloaded files
local_directory = os.path.expanduser("~/Desktop")

# List all files in the local directory
files_in_directory = os.listdir(local_directory)

# Define the substring you want to filter for
substring = "Flag Description"

# Filter files to include only CSV files that contain the specified substring
csv_files = [file for file in files_in_directory if file.endswith(".csv") and substring in file]

# If you want to find the latest file based on modification time, you can use this:
latest_file = max(csv_files, key=lambda x: os.path.getmtime(os.path.join(local_directory, x)))

# Construct the full path to the latest file
latest_file_path = os.path.join(local_directory, latest_file)

# Read the latest CSV file into a Pandas DataFrame
df_1 = pd.read_csv(latest_file_path, on_bad_lines='warn')

# Update the column name
df_1.rename(columns={'﻿ChannelAdvisor Order ID': 'Channel Advisor Order ID'}, inplace=True)

# Convert the 'Flag Description' column to a string data type
df_1['Flag Description'] = df_1['Flag Description'].astype('str')

# Now, you can work with the data in the DataFrame 'df_1'
# For example, you can print the first few rows of the DataFrame:
print(df_1.head())

# Create a public url 
# https://docs.google.com/spreadsheets/d/1NuX2xd9X09qbe5TgwTPoena5VHxwt1b08osy0T7ZijE/edit#gid=1767860156

# Get spreadsheets key from url
gsheetkey = "1NuX2xd9X09qbe5TgwTPoena5VHxwt1b08osy0T7ZijE"

# Sheet name to be read
sheet_name = 'Open_Issues'

url=f'https://docs.google.com/spreadsheet/ccc?key={gsheetkey}&output=xlsx'
df_2 = pd.read_excel(url,sheet_name=sheet_name)
print(df_2.head())
print(df_2.dtypes)

# Create a set of strings to check for
strings_to_check = {'On Issues File', 'On Issue File', 'Issue file', 'issues file', 'issue File',
                    'On Issues Fike', 'On Issue file', 'on issue file', 'on issues file', 'On Issues Files',
                    'On Issue Files', 'Issue files', 'issues files', 'issue File', 'On Issue files', 'on issue files',
                    'on issues files'}

# Create a new column 'Add Note' in df_1
df_1['Add Note'] = ''

# Iterate through the rows of df_1
for index, row in df_1.iterrows():
    value = row['Channel Advisor Order ID']
    flag_description = row['Flag Description']

    # Check if the value exists in column 'PO#' of df_2 and if any of the strings in the list are not in flag_description
    if value in df_2['PO#'].values and all(string not in flag_description for string in strings_to_check):
        df_1.at[index, 'Add Note'] = 'Needs On Issue File Note'
    else:
        df_1.at[index, 'Add Note'] = ''
   
# Generate the current date and time as a string
current_datetime = datetime.now().strftime("%Y-%m-%d")

# Define the output file name with the date and time
output_file_name = f"CA Flag Description Issue Update_{current_datetime}.csv"

# Export the transformed data to a new CSV file in the Desktop
file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
df_1.to_csv(file_path, index=False)

# Print data frame save in df_1
print(df_1.head())
print(df_1.dtypes)
































