from ftplib import FTP
import pandas as pd
import xlwings as xw
import os
from dotenv import load_dotenv
from datetime import datetime
import datetime

# Load the .env file
load_dotenv()

# Usage credentials
local_folder_path = os.path.expanduser("~/Desktop")  # Specify folder path save
remote_folder = '/OSA/SDC/WESTIN_BCTC'  # Specify folder path
# substring = 'SDC_BCTC_ACES'  # Specify the substring to filter files

# Get the variables
ftp_uname = os.getenv('ftpUname')
ftp_pass = os.getenv('ftpPass')
ftp_host = os.getenv('ftpHost')


def download_files_from_ftp(local_folder_path, ftp_host, ftp_uname, ftp_pass, remote_working_directory, file_list):
    try:
        # Create an FTP client instance
        ftp = FTP(timeout=60)

        # Connect to the FTP server
        ftp.connect(ftp_host)

        # Login to the FTP server
        ftp.login(ftp_uname, ftp_pass)
        print("Successfully connected and logged in to the FTP server.")

        # Change current working directory if specified
        if not (remote_working_directory == None or remote_working_directory.strip() == ""):
            _ = ftp.cwd(remote_working_directory)

        # Loop over the list of files
        for file_name in file_list:
            # Derive the local file path by appending the local folder path with the remote filename
            local_file_path = os.path.join(local_folder_path, file_name)

            print(f"Downloading file {file_name}")

            # Download FTP file using retrbinary function
            with open(local_file_path, "wb") as file:
                ftp.retrbinary(f"RETR {file_name}", file.write)

            # Read the downloaded file into a pandas DataFrame
            try:
                df = pd.read_csv(local_file_path, sep="|", low_memory=False)

                # Remove ".txt" extension from filename
                file_name = os.path.splitext(file_name)[0]

                # Save DataFrame to Excel for each file
                excel_file_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{file_name}.xlsx")
                df.to_excel(excel_file_path, index=False, freeze_panes=(1, 0))

                print(f'Saved DataFrame to {excel_file_path}')  # prints the file path

                # Open the Excel file and set all columns width to 15
                with xw.App(visible=False) as app:
                    wb = xw.Book(excel_file_path)

                    # Loop through all worksheets in the workbook
                    for ws in wb.sheets:
                        # Loop through all columns in the worksheet
                        for column in ws.api.UsedRange.Columns:
                            column.ColumnWidth = 15

                    # Save the workbook if needed
                    wb.save()

                    # Close the workbook
                    wb.close()

                print(df.head())  # prints the first 5 rows of the DataFrame

            except Exception as e:
                print(f"An error occurred while reading the file {file_name}: {str(e)}")


        # Send QUIT command to the FTP server and close the connection
        ftp.quit()

        return file_list

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None


# Call the function with the provided parameters
file_list = ["WestinAutomotive_Description_C1.txt", "WestinAutomotive_Attributes_F1.txt",
             "WestinAutomotive_EXPI_E1.txt",
             "WestinAutomotive_Item_Segment_B1.txt", "WestinAutomotive_Packaging_H1.txt",
             "WestinAutomotive_Pricing_D1.txt"]
downloaded_files = download_files_from_ftp(local_folder_path, ftp_host, ftp_uname, ftp_pass, remote_folder, file_list)


def read_excel_1():
    global df1 # Declare df as global to update it
    # Specify the path to your local directory where you downloaded files
    local_directory = os.path.expanduser("~/Desktop")

    # List all files in the local directory
    files_in_directory = os.listdir(local_directory)

    # Define the substring you want to filter for
    substring = "WestinAutomotive_Description_C1.xlsx"

    # Filter files to include only CSV files that contain the specified substring
    excel_files = [file for file in files_in_directory if file.endswith(".xlsx") and substring in file]

    # If you want to find the latest file based on modification time, you can use this:
    latest_file = max(excel_files, key=lambda x: os.path.getmtime(os.path.join(local_directory, x)))

    # Construct the full path to the latest file
    latest_file_path = os.path.join(local_directory, latest_file)

    try:
       if latest_file_path[-5:] == ".xlsx":
          df1 = pd.read_excel(latest_file_path)

       # Make the second row as the new header
       df1.columns = df1.iloc[0]

       # Remove the first row
       df1= df1.iloc[1:]

       # List of columns to drop
       columns_to_drop = ['Date / Time', 'File ID', 'Maint Type', 'Language Code', 'Description Sequence']

       # Drop the columns
       df1.drop(columns=columns_to_drop, inplace=True)

       # Strip leading and trailing spaces from column names
       df1.columns = df1.columns.str.strip()

       # Print the dataframe columns
       print(df1.columns)

       # Pivot the DataFrame
       if 'Part Number' in df1.columns:
           df1 = df1.pivot(index=['Part Number', 'Brand AAIA ID'], columns='Description Code',
                                      values='Description')

           # Reset the index
           df1.reset_index(inplace=True)

           # Print the transformed DataFrame
           print(f'test{df1.head(3)}')

       else:
           print("Error: 'Part Number' column is missing in the DataFrame.")

       # Generate the current date and time as a string
       current_datetime = datetime.datetime.now().strftime("%Y-%m-%d")

       # Define the output file name with the date and time
       output_file_name = f"WestinAutomotive_Description_C1_{current_datetime}.xlsx"

       # Export the filtered data to a new Excel file
       file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
       df1.to_excel(file_path, index=False, freeze_panes=(1, 0))

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
        print("Error: The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        print(f"Error: No such file as {latest_file_path}")
        return None
    
def read_excel_2():
    global df2 # Declare df as global to update it
    # Specify the path to your local directory where you downloaded files
    local_directory = os.path.expanduser("~/Desktop")

    # List all files in the local directory
    files_in_directory = os.listdir(local_directory)

    # Define the substring you want to filter for
    substring = "WestinAutomotive_Attributes_F1.xlsx"

    # Filter files to include only CSV files that contain the specified substring
    excel_files = [file for file in files_in_directory if file.endswith(".xlsx") and substring in file]

    # If you want to find the latest file based on modification time, you can use this:
    latest_file = max(excel_files, key=lambda x: os.path.getmtime(os.path.join(local_directory, x)))

    # Construct the full path to the latest file
    latest_file_path = os.path.join(local_directory, latest_file)

    try:
       if latest_file_path[-5:] == ".xlsx":
          df2 = pd.read_excel(latest_file_path)

       # Make the second row as the new header
       df2.columns = df2.iloc[0]

       # Remove the first row
       df2= df2.iloc[1:]

       # List of columns to drop
       columns_to_drop = ['Date / Time', 'File ID', 'Maint Type', 'Brand AAIA ID', 'PADB Attribute',
                          'Attribute UOM', 'PADB StyleID', 'Record Sequence', 'Multi Value Quantity',
                          'Multi Value Sequence', 'Language Code']

       # Drop the columns
       df2.drop(columns=columns_to_drop, inplace=True)

       # Strip leading and trailing spaces from column names
       df2.columns = df2.columns.str.strip()

       # Print the dataframe columns
       print(df2.columns)

       # Pivot the DataFrame
       if 'Part Number' in df2.columns:
           df2 = df2.pivot(index=['Part Number'], columns='Attribute ID (Type)', values='Attribute Data')
                                      
           # Reset the index
           df2.reset_index(inplace=True)

           # Print the transformed DataFrame
           print(f'test{df2.head(3)}')

       else:
           print("Error: 'Part Number' column is missing in the DataFrame.")

       # Generate the current date and time as a string
       current_datetime = datetime.datetime.now().strftime("%Y-%m-%d")

       # Define the output file name with the date and time
       output_file_name = f"WestinAutomotive_Attributes_F1_{current_datetime}.xlsx"

       # Export the filtered data to a new Excel file
       file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
       df2.to_excel(file_path, index=False, freeze_panes=(1, 0))

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
        print("Error: The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        print(f"Error: No such file as {latest_file_path}")
        return None
    

def read_excel_3():
    global df3 # Declare df as global to update it
    # Specify the path to your local directory where you downloaded files
    local_directory = os.path.expanduser("~/Desktop")

    # List all files in the local directory
    files_in_directory = os.listdir(local_directory)

    # Define the substring you want to filter for
    substring = "WestinAutomotive_Item_Segment_B1.xlsx"

    # Filter files to include only CSV files that contain the specified substring
    excel_files = [file for file in files_in_directory if file.endswith(".xlsx") and substring in file]

    # If you want to find the latest file based on modification time, you can use this:
    latest_file = max(excel_files, key=lambda x: os.path.getmtime(os.path.join(local_directory, x)))

    # Construct the full path to the latest file
    latest_file_path = os.path.join(local_directory, latest_file)

    try:
       if latest_file_path[-5:] == ".xlsx":
          df3 = pd.read_excel(latest_file_path)

       # Make the second row as the new header
       df3.columns = df3.iloc[0]

       # Remove the first row
       df3 = df3.iloc[1:]

       # List of columns to drop
       columns_to_drop = ['Date / Time', 'File ID', 'Maint Type', 'Brand AAIA ID', 'Hazardous Material Code (Y/N)',
                          'Base Item Number', 'Item-Level GTIN (UPC) Qualifier', 'Brand Label', 'SubBrand AAIA ID',
                          'SubBrand Label', 'Item Quantity Size', 'Item Quantity Size UOM', 'Container Type', 
                          'Quantity per Application Qualifier', 'Quantity per Application', 'Quantity per Application UOM',
                          'Item-Level Effective Date ', 'Available Date ', 'Minimum Order Quantity', 'Minimum Order Quantity UOM',
                          'Product Group Code', 'Product Sub-Group Code', 'Product Category Code', 'UNSPSC Code',
                          'VMRS Code (Heavy Duty)' ]

       # Drop the columns
       df3.drop(columns=columns_to_drop, inplace=True)

       # Strip leading and trailing spaces from column names
       df3.columns = df3.columns.str.strip()

       # Print the dataframe columns
       print(df3.columns)

       # Generate the current date and time as a string
       current_datetime = datetime.datetime.now().strftime("%Y-%m-%d")

       # Define the output file name with the date and time
       output_file_name = f"WestinAutomotive_Item_Segment_B1_{current_datetime}.xlsx"

       # Export the filtered data to a new Excel file
       file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
       df3.to_excel(file_path, index=False, freeze_panes=(1, 0))

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
        print("Error: The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        print(f"Error: No such file as {latest_file_path}")
        return None



def read_excel_4():
    global df4 # Declare df as global to update it
    # Specify the path to your local directory where you downloaded files
    local_directory = os.path.expanduser("~/Desktop")

    # List all files in the local directory
    files_in_directory = os.listdir(local_directory)

    # Define the substring you want to filter for
    substring = "WestinAutomotive_Packaging_H1.xlsx"

    # Filter files to include only CSV files that contain the specified substring
    excel_files = [file for file in files_in_directory if file.endswith(".xlsx") and substring in file]

    # If you want to find the latest file based on modification time, you can use this:
    latest_file = max(excel_files, key=lambda x: os.path.getmtime(os.path.join(local_directory, x)))

    # Construct the full path to the latest file
    latest_file_path = os.path.join(local_directory, latest_file)

    try:
       if latest_file_path[-5:] == ".xlsx":
          df4 = pd.read_excel(latest_file_path)

       # Make the second row as the new header
       df4.columns = df4.iloc[0]

       # Remove the first row
       df4 = df4.iloc[1:]

       # List of columns to drop
       columns_to_drop = ['Date / Time', 'File ID', 'Maint Type', 'Brand AAIA ID', 'Package Level GTIN ', 'Electronic Product Code',
                          'Package Bar Code Characters ', 'Package UOM', 'Quantity of Eaches in Package', 'Inner Quantity',
                          'Inner Quantity UOM', 'Orderable Package', 'UOM for Dimensions', 'UOM for Weight', 'Weight Variance (%)',
                          'Dimensional Weight', 'Stacking Factor']

       # Drop the columns
       df4.drop(columns=columns_to_drop, inplace=True)

       # Strip leading and trailing spaces from column names
       df4.columns = df4.columns.str.strip()

       # Print the dataframe columns
       print(df4.columns)

       # Generate the current date and time as a string
       current_datetime = datetime.datetime.now().strftime("%Y-%m-%d")

       # Define the output file name with the date and time
       output_file_name = f"WestinAutomotive_Packaging_H1_{current_datetime}.xlsx"

       # Export the filtered data to a new Excel file
       file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
       df4.to_excel(file_path, index=False, freeze_panes=(1, 0))

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
        print("Error: The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        print(f"Error: No such file as {latest_file_path}")
        return None
    

def read_excel_5():
    global df5 # Declare df as global to update it
    # Specify the path to your local directory where you downloaded files
    local_directory = os.path.expanduser("~/Desktop")

    # List all files in the local directory
    files_in_directory = os.listdir(local_directory)

    # Define the substring you want to filter for
    substring = "WestinAutomotive_Pricing_D1.xlsx"

    # Filter files to include only CSV files that contain the specified substring
    excel_files = [file for file in files_in_directory if file.endswith(".xlsx") and substring in file]

    # If you want to find the latest file based on modification time, you can use this:
    latest_file = max(excel_files, key=lambda x: os.path.getmtime(os.path.join(local_directory, x)))

    # Construct the full path to the latest file
    latest_file_path = os.path.join(local_directory, latest_file)

    try:
       if latest_file_path[-5:] == ".xlsx":
          df5 = pd.read_excel(latest_file_path)

       # Make the second row as the new header
       df5.columns = df5.iloc[0]

       # Remove the first row
       df5 = df5.iloc[1:]

       # List of columns to drop
       columns_to_drop = ['Date / Time', 'File ID', 'Maint Type', 'Brand AAIA ID', 'Price Sheet Number', 'Currency Code', 'Price Sheet Level Effective Date',
                          'Expiration Date', 'Price UOM', 'Price Break Quantity', 'Price Break Quantity UOM']

       # Drop the columns
       df5.drop(columns=columns_to_drop, inplace=True)

       # Strip leading and trailing spaces from column names
       df5.columns = df5.columns.str.strip()

       # Print the dataframe columns
       print(df5.columns)

       # Pivot the DataFrame
       if 'Part Number' in df5.columns:
           df5 = df5.pivot(index=['Part Number'], columns='Price Type',
                                      values='Price')

           # Reset the index
           df5.reset_index(inplace=True)

           # Print the transformed DataFrame
           print(f'test{df5.head(3)}')

       else:
           print("Error: 'Part Number' column is missing in the DataFrame.")

       # Generate the current date and time as a string
       current_datetime = datetime.datetime.now().strftime("%Y-%m-%d")

       # Define the output file name with the date and time
       output_file_name = f"WestinAutomotive_Pricing_D1_{current_datetime}.xlsx"

       # Export the filtered data to a new Excel file
       file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
       df5.to_excel(file_path, index=False, freeze_panes=(1, 0))

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
        print("Error: The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        print(f"Error: No such file as {latest_file_path}")
        return None

   
def merge_1(df1, df2, df3, df4, df5):
      
    merged_df = pd.merge(df1, df2, on='Part Number', how='left')
    merged_df = pd.merge(merged_df, df3, on='Part Number', how='left')
    merged_df = pd.merge(merged_df, df4, on='Part Number', how='left')
    merged_df = pd.merge(merged_df, df5, on='Part Number', how='left')

    merged_df['JBR'] = pd.to_numeric(merged_df['JBR'], errors='coerce').round(2)
    merged_df['RET'] = pd.to_numeric(merged_df['RET'], errors='coerce').round(2)
    merged_df['RMP'] = pd.to_numeric(merged_df['RMP'], errors='coerce').round(2)

    print(f'print first 5 row to {merged_df.head(5)}')

    # Generate the current date and time as a string
    current_datetime = datetime.datetime.now().strftime("%Y-%m-%d")

    # Define the output file name with the date and time
    output_file_name = f"WES_Westin_SDC_Master_File_{current_datetime}.xlsx"

    # Export the filtered data to a new Excel file
    file_path = os.path.join(os.path.expanduser("~"), "Desktop", output_file_name)
    merged_df.to_excel(file_path, index=False, freeze_panes=(1, 0))

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


# Call the function with the provided parameters
file_list = ["WestinAutomotive_Description_C1.txt", "WestinAutomotive_Attributes_F1.txt", "WestinAutomotive_EXPI_E1.txt",
             "WestinAutomotive_Item_Segment_B1.txt", "WestinAutomotive_Packaging_H1.txt", "WestinAutomotive_Pricing_D1.txt"
             ]

downloaded_files = download_files_from_ftp(local_folder_path, ftp_host, ftp_uname, ftp_pass, remote_folder, file_list)
   
# Call the function
manipulation_1 = read_excel_1()
manipulation_2 = read_excel_2()
manipulation_3 = read_excel_3()
manipulation_4 = read_excel_4()
manipulation_5 = read_excel_5()
df_merge_1 = merge_1(df1, df2, df3, df4, df5)








