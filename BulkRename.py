import os
import sys
import Debug as d
# make UI to use all these tools
# create debug system
def main():
    d.dPrint(sys.argv)
    path = sys.argv[1]
    d.dPrint(sys.argv[2])
    folderNames = sys.argv[2].split(',')
    d.dPrint(folderNames)
    d.dPrint(folderNames)
    header = sys.argv[3]
    d.dPrint(path)
    
    d.dPrint(os.getcwd())
    for n in folderNames:
        os.chdir(path + n)
        d.dPrint(os.path)
        for f in os.listdir():
            new_name = header + f
            os.rename(f, new_name)
            
main()

