import os
import sys
import Debug as d
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import tkinter as tk

def renameFiles(header, folders):
    # Imperfect check, only checks first element
    if(folders is not None and isinstance(folders, list) and isinstance(folders[0], str)):
        renameFilesList(header=header, folderPaths=folders)
    elif(isinstance(folders, str)):
        renameFilesString(header=header, folderPath=folders)
    else:
        raise Exception("expected folderPaths to be str or list[str]")

def renameFilesString(header, folderPath: str):
    os.chdir(folderPath)
    for path in os.listdir():
        renameFile(header=header, path=path)
       
def renameFilesList(header, folderPaths: list[str]):
    for f in folderPaths:
        renameFiles(header=header, folders = f)

def renameFile(header, path):
    new_name = header + path
    os.rename(path, new_name)

# Might want to use less global variables
def deleteLabel(label, b):
    paths.remove(label.cget("text"))
    global pathIndex
    # This doesn't seem to work all the time, overlapping is occuring
    pathIndex -= 1
    label.destroy()
    b.destroy()
    
#Create an instance of Tkinter frame
win = Tk()
win.title("File Renaming Utility")
#Define the geometry
win.geometry("750x750")
Label(win, text="Rename Header:", font=('Aerial 16 bold')).grid(row=1)

e1 = Entry(win)
e1.grid(row = 1, column = 2)
paths = []
pathIndex = 0

cbValue = tk.StringVar()
sectionComboBox = ttk.Combobox(win, textvariable=cbValue)
sectionComboBox['width'] = 5
sectionComboBox['values'] = ['', '22 00 ', '23 00 ']
sectionComboBox['state'] = 'readonly'
sectionComboBox.grid(row = 1, column = 1)
sectionComboBox.current(0)

def select_folder():
    global paths
    global pathIndex
    paths.append(filedialog.askdirectory(title="Select Folder"))
    # add x button to remove paths
    l = Label(win, text = paths[pathIndex], font = 11)
    l.grid(row = 7 + pathIndex)
    b = ttk.Button(win, text = "Delete", command = lambda: deleteLabel(l, b))
    b.grid(row = 7 + pathIndex, column = 1)
    print(paths[pathIndex])
    pathIndex += 1
    
#Create a label and a Button to Open the dialog
# Label(win, text="Click the Button to Select a Folder", 
#     font = ('Aerial 18 bold')).grid(row = 2, column = 1)
selectButton = ttk.Button(win, text = "Add Folder", command = select_folder)
selectButton.grid(row = 1, column = 4)
renameButton = ttk.Button(win, text = "Rename All",
    # command = renameFiles(header = e1.get(), folderPaths = path))
    command=lambda: renameFiles((sectionComboBox.get() + e1.get()), paths)).grid(row = 1, column = 5)
renameButton
win.mainloop()


