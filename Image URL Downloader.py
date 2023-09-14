import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import os
import requests


# Declare df as a global variable
df = None

# initalise the tkinter GUI
root = tk.Tk()
root.title("Image URL Downloader")

root.geometry("600x600") # set the root dimensions
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.

# Frame for TreeView
frame1 = tk.LabelFrame(root, text="Excel Data")
frame1.place(height=300, width=600)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File")
file_frame.place(height=200, width=600, rely=0.65, relx=0)

# Buttons
button1 = tk.Button(file_frame, text="Browse for File", command=lambda: File_dialog())
button1.place(rely=0.2, relx=0.01)

button2 = tk.Button(file_frame, text="Download URL Image", command=lambda: Transform_1())
button2.place(rely=0.4, relx=0.01)

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)

# Treeview Widget
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

            # Specify the desktop folder where you want to save images
            desktop_folder = os.path.expanduser("~/Desktop/Image URL Downloads")
            
            # Create the folder if it doesn't exist
            if not os.path.exists(desktop_folder):
                os.makedirs(desktop_folder)

            # Iterate through rows in the DataFrame
            for _, row in df.iterrows():
                image_url = row["Product Image 1"]
                custom_image_name = row["Custom Image Name"]
                if pd.notna(image_url):
                    image_extension = os.path.splitext(image_url)[-1].split('?')[0]  # Get the file extension
                    image_name = f"{custom_image_name}{image_extension}"
                    image_path = os.path.join(desktop_folder, image_name)

                    try:
                        # Send an HTTP request to download the image
                        response = requests.get(image_url)
                        response.raise_for_status()

                        # Save the image to the specified folder
                        with open(image_path, "wb") as file:
                            file.write(response.content)

                        print(f"Downloaded: {image_name}")
                    except Exception as e:
                        print(f"Failed to download {image_name}: {e}")

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None

root.mainloop()








 