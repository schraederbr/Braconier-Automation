import os

sections = {
    "22 00 ": ["Winsupply "],
    "23 00 ": ["A&P ", "Allied ", "CFM ", "Engineered Products "
    ,"Finn & Associates ", "Kuck ", "MVA ", "Preferred Water ",
    ]
}

def renameFiles(header, folders):
    # Imperfect check, only checks first element
    if(folders is not None and isinstance(folders, list) and isinstance(folders[0], str)):
        __renameFilesList(header=header, folderPaths=folders)
    elif(isinstance(folders, str)):
        __renameFilesString(header=header, folderPath=folders)
    else:
        raise Exception("expected folderPaths to be str or list[str]")

def __renameFilesString(header, folderPath: str):
    os.chdir(folderPath)
    for path in os.listdir():
        __renameFile(header=header, path=path)
    
def __renameFilesList(header, folderPaths: list[str]):
    for f in folderPaths:
        renameFiles(header=header, folders = f)

def __renameFile(header, path):
    new_name = header + path
    os.rename(path, new_name)




