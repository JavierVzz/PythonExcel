# Javier Vazquez
# Python 3.5.2

import cellInverterClass

def main():
    invert = cellInverterClass.cellInverter()
    print(invert.displayExcelFiles())
    file = input("File: ")
    wb, ws = invert.readExcelFile(file)
    invert.displayContent()
    invert.invert()

if __name__ == '__main__':
    main()