import os
import tkinter
from tkinter import filedialog as fd
import openpyxl
import sys
import shutil

def getNames():
    # Let user select a spreadsheet file
    filename = fd.askopenfilename()
    if not filename:
        sys.exit(0)

    # Open the spreadsheet using openpyxl
    wb = openpyxl.load_workbook(filename)

    # Get the first sheet
    sheet = wb.active

    # Get the first column
    column = sheet['A']

    # Loop through the column and put the values in a list
    names = []
    for cell in column:
        names.append(cell.value)
    
    # Close the spreadsheet
    wb.close()

    # Print the names
    return names

def getImageFolder():
    # Tell user to select a folder with images
    folder = fd.askdirectory()
    if not folder:
        sys.exit(0)
    
    # Return the folder directory path
    return folder

def main():
    # Get the names
    root = tkinter.Tk()
    root.withdraw()
    # Tell the user to select a spreadsheet with message
    tkinter.messagebox.showinfo('Select a spreadsheet', 'Select a spreadsheet')
    names = getNames()
    print(names)

    # Get the folder's name
    # Tell the user to select a folder with images
    tkinter.messagebox.showinfo('Select a folder with images', 'Select a folder with images')
    folder = getImageFolder()
    print(folder)

    # Create new folder in one directory up to hold the renamed images (if it exists, delete it)
    newFolder = os.path.join(os.path.dirname(folder), 'renamedImages')
    if os.path.exists(newFolder):
        shutil.rmtree(newFolder)
    os.mkdir(newFolder)

    # Loop through the names
    for name in names:
        # Look for a file with the name (case insensitive) without an extension and make sure it's an image file
        for i, filename in enumerate(os.listdir(folder)):
            if filename.lower().startswith(name.lower()):
                #print('{:03d} {}'.format(i+1, name))
                # Copy the file to the new folder with the new name (with the same extension), keeping original file intact
                shutil.copy(os.path.join(folder, filename), os.path.join(newFolder, '{:03d} {}.{}'.format(i+1, name, filename.split('.')[-1])))
                break

if __name__ == "__main__":
    main()