import os

#Not sure if I should have this here, 
# maybe move other dictionaries from MaterialsChart.py here
sections = {
    "22 00 ": ["Winsupply"],
    "23 00 ": ["A&P", "Allied", "CFM", "Engineered Products"
    ,"Finn & Associates", "Kuck", "MVA", "Preferred Water"
    ]
}

#may want to rename this function, to somethng like renameFileFolder
def renameFiles(folders, header='', toRemove = False, charsToRemove = 0):
    # Imperfect check, only checks first element
    if(folders is not None and isinstance(folders, list) and isinstance(folders[0], str)):
        if toRemove:
            removeCharactersList(folders, charsToRemove)
        __renameFilesList(header=header, folderPaths=folders)
    elif(isinstance(folders, str)):
        if toRemove:
            removeCharactersString(folders, charsToRemove)
        __renameFilesString(header=header, folderPath=folders)
    else:
        raise Exception("expected folderPaths to be str or list[str]")



def removeCharactersString(path, charsToRemove = 0):
    os.chdir(path)
    for path in os.listdir():
        print(path)
        print(path[charsToRemove:])
        os.rename(path, path[charsToRemove:])

def removeCharactersList(paths, charsToRemove = 0):
    for f in paths:
        removeCharactersString(f, charsToRemove)
#make sure this __ usage is correct
# if these functions don't change, possible combine them 
#, if it makes them more readable
def __renameFilesString(header, folderPath):
    os.chdir(folderPath)
    for path in os.listdir():
        __renameFile(path=path, header=header)
    
#This seems a little off, double check
def __renameFilesList(header, folderPaths):
    for f in folderPaths:
        renameFiles(header=header, folders = f)

def __renameFile(header, path):
    new_name = header + path
    os.rename(path, new_name)




