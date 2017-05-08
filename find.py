import sys
import xlrd

filename = sys.argv[1]
sheetnumber = (1, 2)
print filename

for name, sheet_name in zip(filename, sheetnumber):
    print name
    book = xlrd.open_workbook(name)
    sheet = book.sheet_by_name(sheet_name)
    for row in range(sheet.nrows):
        for column in range(sheet.ncols):
            if sheet.cell(row,column).value == value:
                print row, column
