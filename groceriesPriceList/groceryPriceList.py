# Javier Vazquez

import openpyxl, os, re, pprint, sys, logging
from groceriesPriceList.readExcel import excelOperations

def main():
    eo = excelOperations()
    pprint.pprint(eo.displayExcelFiles())
    file = input("File: ")
    eo.readExcelFile(file)
    eo.createDictionary()
    eo.saveDictionary("w")
    eo.loadDictionary()
    eo.printDictionary()
    key = input("Produce: ")
    value = input("Price: ")
    updateKey = eo.updateKey(key, value)
    if updateKey == True:
        print(key + " has a new price: "+ value)
        eo.updateExcel(key, value)
        input("Press ENTER!!")
        eo.printDictionary()
    else:
        print(key + " does not exist!!!")
        input("Press ENTER!!")


    print("Done")

main()