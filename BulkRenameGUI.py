#could probably clean up these imports
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from RenameTools import *
import warnings
import shutil

# add buttons gray out when they shouldn't be used
# like if 

def deleteLabel(label, b):
    '''Deletes tkinter label and button. Removes associated path from paths'''
    # Glitches when deleting then adding again, seems like tkinter problem, not an index problem 
    paths.remove(label.cget("text"))
    label.destroy()
    b.destroy()
    
#Sets up main tkinter window
win = Tk()
win.title("File Renaming Utility")
# add autoresize if possible
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
        l.grid(row = 7 + len(paths), columnspan=4)
        #button to delete folder path
        b = ttk.Button(win, text = "Remove", command = lambda: deleteLabel(l, b))
        b.grid(row = 7 + len(paths), column = 5)
    
addFolderButton = ttk.Button(win, text = "Add Folder", command = selectFolder)
addFolderButton.grid(row = 1, column = 4)

renameButton = ttk.Button(win, text = "Rename All Files",
    command=lambda: renameFiles((sectionComboBox.get() + e1.get()), paths))
renameButton.grid(row = 1, column = 5)


#Automatically open relevant files to edit, like Cover page and warranty
def autoRenameFolders(folders):
    '''Renames folders based on known section and manufacturers'''
    for f in folders:
        autoRenameFolder(f)


def autoRenameFolder(f):
    '''Renames single folder based on knonwn sections and manufacturers'''
    bookPath = ""
    validPaths = []
    if not(f is None):
        os.chdir(f)
        validPaths.append(f)
        for p in os.listdir():
            if(os.path.isdir(p)):
                os.chdir(p)
                # might be better way to avoid using recognized variable
                recognized = False
                for key in sections:
                    for s in sections[key]:
                        if(s in p):
                            renameFiles(key + s, os.getcwd())
                            recognized = True         
                if not(recognized):
                    if ("Book Assembly" in p):
                        bookPath = os.getcwd()
                    elif not ("Book Assembly" in p):
                        warnings.warn("Folder name not recognized, defaulting to 23 00 + FOLDERNAME") 
                        renameFiles("23 00 " + p, os.getcwd())
                elif(recognized):
                    validPaths.append(os.getcwd())
                os.chdir(f)
    if(bookPath == ""):
        warnings.warn("Book Assembly folder not found, files will not be copied")
    else:
        # Add Auto unpack ZIP files
        # Add stitching the documents together, hopefully with bluebeam
        for p in validPaths:
            os.chdir(p)
            for f in os.listdir():
                if(os.path.isfile(f)):
                    shutil.copy(p+"\\"+ f, bookPath+"\\"+f)

        pass

autoButton = ttk.Button(win, text = "Closeout Mode", command = lambda: autoRenameFolders(paths))
autoButton.grid(row = 2, column = 4)

win.mainloop()



