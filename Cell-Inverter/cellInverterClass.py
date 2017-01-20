# Javier Vazquez
# Python 3.5.2

import readExcel, os, openpyxl

class cellInverter(readExcel.excelOperations):

    def __init__(self):
        super().__init__()

    def invert(self, file= "inverted.xlsx"):
        wbSource, wsSource = super().wb_ws
        wbDest = openpyxl.Workbook()
        wsDest = wbDest.active
        # print(wbSource, wsSource)
        for i in range(1, wsSource.max_row + 1):
            for j in range(1, wsSource.max_column + 1):
                print(wsSource.cell(row=i, column=j).value)
                wsDest.cell(row=j, column=i).value = wsSource.cell(row=i, column=j).value
        wbDest.save(file)


if __name__ == '__main__':
    print("Direct access to "+ os.path.basename(__file__))
else:
    print(os.path.basename(__file__)+" class instance")