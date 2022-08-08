from pathlib import Path

def verifyJobNumber(jNumber):
    if(type(jNumber) == str):
        if(jNumber.isdigit()):
            if(len(jNumber) == 4):
                return True
    print("Invalid job number")
    return False

#Need to make this work with any year. Probably should just use regex like this: M:\BPH Bids FY####
def getXlFiles(jobNumbers, extensions=[".xls", ".xlsx"]):
    print(jobNumbers)
    xlFiles = []
    jobPaths = []
    #May want to do regex on .rglob(*) if it's faster
    #jobPaths seems to have duplicates based on how many jobs inputted. This isn't a problem
    #because in filecache, no duplicates allowed. 
    print(Path("M:\\"))
    for p in Path("M:\\").glob("BPH Bids FY*"):
        print(p)
        for j in jobNumbers:
            if verifyJobNumber(j):
                for p2 in Path(p).glob("*" + j + "*"):
                    for path in Path(p2).rglob("*BUYOUT*"):
                        jobPaths.append(path)
    print(*jobPaths, sep='\n')
    for folder in jobPaths: 
        print(folder.as_posix())
        #This is slightly impressice but probably good enough
        if("ESTIMATING" in folder.as_posix().upper()):
            for extension in extensions:
                for path in Path(folder).rglob('**/*' + extension):
                    xlFiles.append(path)
    return xlFiles
    
    #maybe return the whole path object instead of just the path str
    
        
    
    
    
    
    