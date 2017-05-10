from openpyxl import Workbook
from openpyxl import load_workbook
import xlrd
import sys
from itertools import izip_longest
from tabulate import tabulate

final = []

filename_one = sys.argv[1]
filename_two = sys.argv[2]

file_one = xlrd.open_workbook(filename_one)
file_two = xlrd.open_workbook(filename_two)

# get all sheets
sheets_no_one = file_one.nsheets
sheets_no_two = file_two.nsheets
sheets_one = file_one.sheet_names()
sheets_two = file_two.sheet_names()

print sheets_one, sheets_two

def log_comparison(error, position, file1_val, file2_val):
    comparison = {}
    comparison['ERROR'] = error
    comparison['POSITION'] = position
    comparison[filename_one.upper()] = file1_val
    comparison[filename_two.upper()] = file2_val
    final.append(comparison)
    print comparison

def compare_sheet(sheet_name, sheet1, sheet2):
    for rownum in range(sheet1.nrows):
        row1 = sheet1.row_values(rownum)
        try:
            row2 = sheet2.row_values(rownum)
        except Exception:
            log_comparison('Missing Row', 'Row No. : %s' % str(rownum+1), '[exists]', '--')
            continue
        for colnum, (c1, c2) in enumerate(izip_longest(row1, row2)):
            if c1 != '':
                if c1 != c2:
                    log_comparison('Unequal Value', 'Row %d Col %d' % (rownum+1, colnum+1), c1, c2)

def display_comparison():
    print tabulate(final, headers='keys', tablefmt = 'grid')


def start():
    print '-'*20
    for a_sheet in sheets_one:
        sheet1 = file_one.sheet_by_name(a_sheet)
        try:
            sheet2 = file_two.sheet_by_name(a_sheet)
        except Exception:
            log_comparison('Missing Sheet', a_sheet, '[exists]', '--')
            continue
        compare_sheet(a_sheet, sheet1, sheet2)
    print 'comparing'

start()
display_comparison()
