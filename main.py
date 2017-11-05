import sys
import xlrd
import xlwt
from xlutils.copy import copy
import time

#first_source_filepath = input("First file path: ")
first_source_filepath = sys.argv[1]

#second_source_filepath = input("Second file path: ")
second_source_filepath = sys.argv[2]

#output_filepath = input("Output file path: ")
output_filepath = sys.argv[3]

start = time.time() 

# Read the first input file
first_source_book = xlrd.open_workbook(first_source_filepath)
first_source_book_first_sheet = first_source_book.sheet_by_index(0) # get first sheet
first_source_book_first_sheet_max_rows = first_source_book_first_sheet.nrows
values = first_source_book_first_sheet.col_values(0, 1)
#print("[+] File: " + first_source_filepath + ", values column A, skipping the header:")
#for value in values:
#    print(value)

# Create a destination file
output_book = copy(first_source_book)
output_book_first_sheet = output_book.get_sheet(0)

# Read the second source file
second_source_book = xlrd.open_workbook(second_source_filepath)
second_source_book_first_sheet = second_source_book.sheet_by_index(0)
values = second_source_book_first_sheet.col_values(0, 1)
#print("[+] File: " + second_source_filepath + ", values column A, skipping the header:")

# Append the line to the destination file:
second_source_book_first_sheet_max_rows = second_source_book_first_sheet.nrows
second_source_book_first_sheet_max_columns = second_source_book_first_sheet.ncols
for row_index in range(1, second_source_book_first_sheet_max_rows): # skip the header this time
#    print("\n[!] Printing row index = " + str(row_index))
    for column_index in range(0, 11): # Columna A-K
        cell_value = second_source_book_first_sheet.cell(row_index, column_index).value
        output_book_first_sheet.write(first_source_book_first_sheet_max_rows+row_index-1,column_index,cell_value)
    for column_index in range(16, second_source_book_first_sheet_max_columns-1): # Columns Q-Z. Do not include last column either.
        cell_value = second_source_book_first_sheet.cell(row_index, column_index).value
        output_book_first_sheet.write(first_source_book_first_sheet_max_rows+row_index-1,column_index-5,cell_value) # Account for the diff Q-Z

output_book.save(output_filepath)
output_book_reopen = xlrd.open_workbook(output_filepath)
output_book_reopen_first_sheet = output_book_reopen.sheet_by_index(0)
#print("output_book_reopen_first_sheet.nrows = " + str(output_book_reopen_first_sheet.nrows))
values = output_book_reopen_first_sheet.col_values(0, 1)
#print("[+] File: " + output_filepath + ", values column A, skipping the header:")
#for value in values:
    #print(value)

print("\n[+] Completed in {0:.5f}".format(time.time() - start) + " seconds")