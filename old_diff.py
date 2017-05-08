from openpyxl import Workbook
from openpyxl import load_workbook
import xlrd
import sys
from itertools import izip_longest

final = {}

rb1 = xlrd.open_workbook(sys.argv[1])
rb2 = xlrd.open_workbook(sys.argv[2])

sheet1 = rb1.sheet_by_index(0)
sheet2 = rb2.sheet_by_index(0)

print sheet1, sheet2
print sheet1.nrows, sheet2.nrows, '---number of rows---\n\n'

print sheet1.row_values(0), '--> row values of 0th row -- \n \n'

for rownum in range(max(sheet1.nrows, sheet2.nrows)):
    print rownum, sheet1.nrows
    if rownum < sheet1.nrows:
        try:
            row_rb1 = sheet1.row_values(rownum)
            row_rb2 = sheet2.row_values(rownum)
        except:
            print 'no such column in sheet 2'
        print row_rb1, row_rb2
        for colnum, (c1, c2) in enumerate(izip_longest(row_rb1, row_rb2)):
            print c1, c2
            if c1 != c2:
                print "Row {} Col {} - {} != {}".format(rownum+1, colnum+1, c1, c2)
    else:
        print 'else k bahar'
        print "Row {} missing".format(rownum+1)

print '______________'
exit()









one = xlrd.open_workbook(sys.argv[1])
two = xlrd.open_workbook(sys.argv[2])

sheet1 = f1.sheet_by_index(0)

print sheet1.nrows
exit()

wb = load_workbook('test.xlsx')

print wb.get_sheet_names()
print wb.sheet_names()
