#Javier Vazquez

import readExcel, os, openpyxl

class insertBlankRows(readExcel.excelOperations):

    def __init__(self):
        super().__init__()

    def insertBlankRows(self, inRow, many, file= "final.xlsx"):
        wbSource, wsSource = super().wb_ws
        wbDest = openpyxl.Workbook()
        wsDest = wbDest.active
        print(wbSource, wsSource)
        for i in range(1, inRow):
            for j in range(1, wsSource.max_column):
                wsDest.cell(row=i, column=j).value = wsSource.cell(row=i, column=j).value
                wsDest.cell(row=i, column=j + wsSource.max_column + 1).value = wsSource.cell(row=i, column=j).value

        for i in range(inRow, inRow + many):
            for j in range(1, wsSource.max_column):
                wsDest.cell(row=i, column=j).value = wsSource.cell(row=i, column=j).value
                wsDest.cell(row=i, column=j + wsSource.max_column + 1).value = " "
        
        for i in range(inRow + many, wsSource.max_row + many + 1):
            for j in range(1, wsSource.max_column):
                wsDest.cell(row=i, column=j).value = wsSource.cell(row= i, column=j).value
                wsDest.cell(row=i, column=j + wsSource.max_column + 1).value = wsSource.cell(row= i - many, column=j).value

        wbDest.save(file)







if __name__ == '__main__':
    print("Direct access to "+ os.path.basename(__file__))
else:
    print(os.path.basename(__file__)+" class instance")
