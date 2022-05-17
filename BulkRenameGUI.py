#could probably clean up these imports
from importlib.resources import path
import re
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import tkinter as tk
from RenameTools import *
import warnings



def deleteLabel(label, b):
    '''Deletes tkinter label and button. Removes associated path from paths'''
    # Glitches when deleting then adding again, seems like tkinter problem, not an index problem 
    paths.remove(label.cget("text"))
    label.destroy()
    b.destroy()
    
#Sets up main tkinter window
win = Tk()
win.title("File Renaming Utility")
win.geometry("750x750")
Label(win, text="Rename Header:", font=('Aerial 16 bold')).grid(row=1)

#Text entry box for filename header
e1 = Entry(win)
e1.grid(row = 1, column = 2)

#Combobox for premade filename headers
sectionComboBox = ttk.Combobox(win)
sectionComboBox['width'] = 5
sectionComboBox['values'] = ['', '22 00 ', '23 00 ']
sectionComboBox['state'] = 'readonly'
sectionComboBox.grid(row = 1, column = 1)
sectionComboBox.current(0)

#Global variable used to hold folder paths
paths = []


def selectFolder():
    '''Saves folder path and creates associated label and delete button'''
    global paths
    path = filedialog.askdirectory(title="Select Folder")
    #Checks if path is not empty (If user cancels folder selection)
    if not(path == ''):
        paths.append(path)
        #label to display folder path
        l = Label(win, text = paths[-1], font = 11)
        l.grid(row = 7 + len(paths))
        #button to delete folder path
        b = ttk.Button(win, text = "Delete", command = lambda: deleteLabel(l, b))
        b.grid(row = 7 + len(paths), column = 1)
    
addFolderButton = ttk.Button(win, text = "Add Folder", command = selectFolder)
addFolderButton.grid(row = 1, column = 4)

renameButton = ttk.Button(win, text = "Rename All",
    command=lambda: renameFiles((sectionComboBox.get() + e1.get()), paths))
renameButton.grid(row = 1, column = 5)

def autoRenameFolders(folders):
    for f in folders:
        autoRenameFolder(f)

def autoRenameFolder(f):
    if not(f is None):
        os.chdir(f)
        for p in os.listdir():
            if(os.path.isdir(p)):
                os.chdir(p)
                # might be better way to avoid using recognized variable
                recognized = False
                # could probably condense these loops
                for s in sections["22 00 "]:
                    if(s in p):
                        renameFiles("22 00 " + s, os.getcwd())
                        recognized = True        
                for s in sections["23 00 "]:
                    if(s in p):
                        renameFiles("23 00 " + s, os.getcwd())
                        recognized = True   
                if not(recognized):
                    if not ("Book Assembly" in p):
                        warnings.warn("Folder name not recognized, defaulting to 23 00 + FOLDERNAME") 
                        renameFiles("23 00 " + p, os.getcwd())
                os.chdir(f)

autoButton = ttk.Button(win, text = "Auto Rename First Folder", command = lambda: autoRenameFolders(paths))
autoButton.grid(row = 2, column = 4)

win.mainloop()



