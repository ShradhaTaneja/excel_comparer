from openpyxl import Workbook
from openpyxl import load_workbook
import xlrd
import sys
from itertools import izip_longest

final = {}

filename_one = sys.argv[1]
filename_two = sys.argv[2]

file_one = xlrd.open_workbook(filename_one)
file_two = xlrd.open_workbook(filename_two)

# get all sheets
sheets_no_one = file_one.nsheets
sheets_no_two = file_two.nsheets
sheets_one = file_one.sheet_names()
sheets_two = file_two.sheet_names()

print '%s -> %s \n %s -> %s' % (filename_one,sheets_no_one,filename_two, sheets_no_two)


def compare(item_type, first_sheet = None, second_sheet = None):
    if item_type == 'number_of_sheets':
        if file_one.nsheets == file_two.nsheets:
            return None
        elif file_one.nsheets > file_two.nsheets:
            return {'high_filename': filename_one, 'low_filename':filename_two, 'high_sheets': sheets_one, 'low_sheets': sheets_two}
        elif file_two.nsheets > file_one.nsheets:
            return {'high_filename': filename_two, 'low_filename':filename_one, 'high_sheets': sheets_two, 'low_sheets': sheets_one}
    if item_type == 'rows':
        return None


def number_of_sheets():
    res = compare('number_of_sheets')
    if res is not None:
        print 'Missing sheets in %s are : %s ' % (res['low_filename'], ', '.join([sheet for sheet in res['high_sheets'] if sheet not in res['low_sheets']]))


def compare_sheet(index, first_file, second_file):
    first_sheet = first_file.sheet_by_index(index)
    second_sheet = second_file.sheet_by_index(index)
    res = compare('rows')
    print first_sheet.row_values(0)


def compare_all_sheets():
    for sheet_idx in range(0, file_one.nsheets):
        print sheet_idx
        compare_sheet(sheet_idx, file_one, file_two)
        exit()

#number_of_sheets()
compare_all_sheets()
#print compare('sheets')
