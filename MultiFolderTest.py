import tkfilebrowser
from tkinter import *

root = Tk()
root.geometry('200x200')
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

dirs = []

def get_directories():
    dirs.append(tkfilebrowser.askopendirnames())
    return dirs

b1 = Button(root, text='select directories...', command=get_directories)
b1.pack()

root.mainloop()