import os
import tkinter as tk
import tkinter.messagebox as messagebox
import sys
import pandas as pd
import shutil
import openpyxl
from tkinter.filedialog import askopenfilename, askdirectory
import importlib

required_modules = ["os", "tkinter", "tkinter.messagebox", "sys", "pandas", "shutil", "openpyxl", "tkinter.filedialog"]

for module in required_modules:
    try:
        importlib.import_module(module)
    except ImportError:
        !pip install {module}

# we don't want a full GUI, so keep the root window from appearing
Tk().withdraw()

# Function to open a file selection dialog and return the selected file's location
def select_file():
  file_location = tkinter.filedialog.askopenfilename(title="Select the model file")

# Function to open a directory selection dialog and return the selected directory's location
def select_directory():
  location = filedialog.askdirectory(title="Select the parent folder location")

# Create a button to open the file selection dialog
file_button = tk.Button(root, text="Select model file", command=select_file)
file_button.pack()

# Create a button to open the directory selection dialog
directory_button = tk.Button(root, text="Select parent folder location", command=select_directory)
directory_button.pack()

# Run the tkinter event loop
root.mainloop()
# Get the user's desktop directory
desktop_dir = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
# Get the selected file and directory locations
file_location = select_file()
location = select_directory()

# Create the parent folder using the prefix and suffix
while True:
    # Get the prefix and suffix from the user
    #//prefix = input("Enter the Part Number: ")      (UNUSED IN THIS CODE)
    suffix = input("Enter the Revision Letter: ")
    # // We can automate rev letters later once we rip data from the part.
    # Get the file name from the file location
    file_name = os.path.basename(file_location)

    # Extract the file name without the extension
    file_name_without_ext = os.path.splitext(file_name)[0]
    # Create the parent folder using the prefix and suffix
    parent_folder = os.path.join(location, f"{file_name_without_ext}_REV_{suffix}")
    #parent_folder = os.path.join(location, f"{prefix}_REV_{suffix}")
    if not os.path.exists(parent_folder):
        os.makedirs(parent_folder)

    # Write the prefix and suffix to a text file
    with open(os.path.join(parent_folder, "prefix_suffix.txt"), "w") as f:
        #f.write(f"Prefix: {prefix}\nSuffix: {suffix}")
        f.write(f"Prefix: {file_name_without_ext}\nSuffix: _REV_{suffix}")

    # Create the "HM" and "Setup" folders inside the parent folder
    hm_folder = os.path.join(parent_folder, "HM")
    setup_folder = os.path.join(parent_folder, "Setup")
    if not os.path.exists(hm_folder):
        os.makedirs(hm_folder)
    if not os:
# Create the "HM" and "Setup" folders inside the parent folder
        hm_folder = os.path.join(parent_folder, "HM")
        setup_folder = os.path.join(parent_folder, "Setup")
    if not os.path.exists(hm_folder):
        os.makedirs(hm_folder)

# Write the file location to the "prefix_suffix.txt" file on a new line
    with open(os.path.join(parent_folder, "prefix_suffix.txt"), "a") as f:
        f.write(f"\nFile location: {file_location}")

# Write the parent folder location to a file on the desktop
    with open(os.path.join(desktop_dir, "parent_folder_location.txt"), "w") as f:
        f.write(f"Parent folder location: {parent_folder}")


# Copy the model file to the parent folder
    shutil.copyfile(file_location, os.path.join(parent_folder, "model.file"))
# Open the parent folder in Explorer
    os.startfile(parent_folder)

# Display a thank you message
    print("Thank you for using this script!")


# Create a CSV file on the desktop with the required information
    with open(os.path.join(desktop_dir, "folder_information.csv"), "w", newline="") as csv_file:
# Write the header row
        csv_file.write("Parent Folder Location,File Location,Suffix,HM Folder Location,Startup\n")

# Write the information for the parent folder
        csv_file.write(f"{parent_folder},{file_location},_REV_{suffix},{hm_folder},true\n")



# Display a message box with the message "All done sorting and cleaning"
    messagebox.showinfo('Data Written, all done')
    #Add tandom Quotes
    print("Your desktop belongs to the bots in the name of botopia")
# Schedule the script to exit after 5 seconds
    messagebox.after(5000, sys.exit)
