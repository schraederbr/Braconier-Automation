from pathlib import Path

#might need more specific filtering to not include irrelevant files, possible name checks, but hopefully something better than tthat
# xlsx and xls should be included. Need to figure out how to interface with both,
# openpyxl and xlrd. openpyxl doesn't work with xls but i think xlrd work with xls.
# Maybe let it work with multiple folderPaths
def verifyJobNumber(jNumber):
    if(type(jNumber) == str):
        if(jNumber.isdigit()):
            if(len(jNumber) == 4):
                return True
    print("Invalid job number")
    return False

#Multiple extensions might have caused a bug with /FileCache error
def getXlFiles(jobNumbers, folderPath = "M:\BPH Bids FY2022", extensions=[".xls", ".xlsx"]):
    print(jobNumbers)
    xlFiles = []
    jobPaths = []
    for j in jobNumbers:
        if verifyJobNumber(j):
            for p in Path(folderPath).glob("*" + j + "*"):
                for path in Path(p).rglob("*Estimating Take-Offs*"):
                    jobPaths.append(path)
    for folder in jobPaths:
        for extension in extensions:
            for path in Path(folder).rglob('**/*' + extension):
                xlFiles.append(path)
        
    return xlFiles
    
    #maybe return the whole path object instead of just the path str
    
        
    
    
    
    
    