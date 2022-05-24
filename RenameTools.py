import os
from typing import List, Union

#Not sure if I should have this here, 
# maybe move other dictionaries from MaterialsChart.py here
sections = {
    "22 00 ": ["Winsupply "],
    "23 00 ": ["A&P ", "Allied ", "CFM ", "Engineered Products "
    ,"Finn & Associates ", "Kuck ", "MVA ", "Preferred Water ",
    ]
}

def renameFiles(header: str, folders: Union[List[str], str]) -> None:
    # Imperfect check, only checks first element
    if(folders is not None and isinstance(folders, list) and isinstance(folders[0], str)):
        __renameFilesList(header=header, folderPaths=folders)
    elif(isinstance(folders, str)):
        __renameFilesString(header=header, folderPath=folders)
    else:
        raise Exception("expected folderPaths to be str or list[str]")

#make sure this __ usage is correct
# if these functions don't change, possible combine them 
#, if it makes them more readable
def __renameFilesString(header: str, folderPath: str) -> None:
    os.chdir(folderPath)
    for path in os.listdir():
        __renameFile(header=header, path=path)
    
def __renameFilesList(header: str, folderPaths: list[str]) -> None:
    for f in folderPaths:
        renameFiles(header=header, folders = f)

def __renameFile(header: str, path: str) -> None:
    new_name = header + path
    os.rename(path, new_name)




