#Javier Vazquez

import BlankRowsClass, os

def main():
    bRows = BlankRowsClass.insertBlankRows()
    print(bRows.displayExcelFiles())
    file = input("File: ")
    wb, ws = bRows.readExcelFile(file)
    #bRows.displayContent()
    inRow = input("Insert blank row(s) at row: ")
    many = input("How many blank row(s) to insert: ")
    bRows.insertBlankRows(int(inRow), int(many))


if __name__ == '__main__':
    main()
