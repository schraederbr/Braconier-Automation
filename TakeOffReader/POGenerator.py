import openpyxl
from tkinter import filedialog
import os
import sqlite3
import subprocess
class LineItem:
    def __init__(self, type, material, name, size, quantity):
        self.type = type
        self.material = material
        self.name = name
        self.size = size
        self.quantity = quantity

    def __str__(self):
        return self.__repr__()
    def __repr__(self):
        return "{},{},{},{},{}\n".format(self.type, self.material, self.name, self.size, self.quantity)

# Query to Combine rows: Not sure what the having count does
# SELECT SUM(quantity) as 'Total Quantity', Type, Material, Name, Size 
# FROM LineItems 
# GROUP BY Size, Name, Material, Type 
# HAVING COUNT(*)>1;


# ID INTEGER PRIMARY KEY AUTOINCREMENT
if __name__ == '__main__':
    con = sqlite3.connect('po.db')
    cur = con.cursor()
    #main lives in the same namespace as top level code interesting, so no need for global calls
    folderPath = filedialog.askdirectory(title="Select Folder With Material Charts")
    #folderPath = "C:\Automation\TakeOffReader\Takeoffs"
    print(cur.execute("DROP TABLE IF EXISTS LineItems").fetchall())
    print(cur.execute("""CREATE TABLE LineItems(
            Type VARCHAR(50) NOT NULL,
            Material VARCHAR(50) NOT NULL,
            Name VARCHAR(50) NOT NULL,
            Size VARCHAR(50) NOT NULL,
            Quantity REAL NOT NULL
            );""").fetchall())
    #print(*(os.listdir(folderPath)), sep='\n')
    for file in os.listdir(folderPath):
        sourcePath  = os.path.join(folderPath, file)
        print(sourcePath)
        source = openpyxl.load_workbook(sourcePath)
        sourceSheet = source.active
        rowIter = sourceSheet.iter_rows(max_row=sourceSheet.max_row, max_col=sourceSheet.max_column)
        currentType = "Unknown"
        currentMaterial = "Unknown"
        currentName = "Unknown"
        lineItems = []
        for i, row in enumerate(rowIter):
            if(sourceSheet["A" + str(i+1)].value is not None):
                currentType = sourceSheet["A" + str(i+1)].value
            if(sourceSheet["G" + str(i+1)].value is not None):
                currentMaterial = sourceSheet["G" + str(i+1)].value
            if(sourceSheet["F" + str(i+1)].value is not None and sourceSheet["F" + str(i+1)].value != "Item"):
                currentName = sourceSheet["F" + str(i+1)].value
            #This if may need to be improved, more detailed, it may have exceptions I haven't covered
            if(sourceSheet["I" + str(i+1)].value is not None and sourceSheet["O" + str(i+1)].value is not None):
                lineItems.append(LineItem(currentType, currentMaterial, currentName, sourceSheet["I" + str(i+1)].value, sourceSheet["O" + str(i+1)].value))   
        fileName = file + ".csv"
        f = open(fileName, "w")
        f.write("{},{},{},{},{}\n".format("Type", "Material", "Name", "Size", "Quantity"))           
        for l in lineItems:
            f.write(str(l))
            query = 'INSERT INTO LineItems(Type, Material, Name, Size, Quantity) VALUES("{}", "{}", "{}", "{}", {})'.format( l.type, l.material, l.name, l.size, l.quantity)
            #print(query)
            cur.execute(query)
    f.close()
    #May want # HAVING COUNT(*)>1; at the end I don't know what this does though
    #Create a separate table with this data in the database:
    print(*cur.execute("""SELECT SUM(quantity) as 'Total Quantity', Type, Material, Name, Size 
    FROM LineItems GROUP BY Size, Name, Material, Type """).fetchall(), sep='\n')
    con.commit()

        
    
    
    
    
    