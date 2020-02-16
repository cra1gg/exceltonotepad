import xlrd
import xlrd.sheet
import csv
import sys
import os

def add_spaces(value, num):
    while len(value) < num:
        value = value + " "
    return value

num_col = int(input("How many columns are there?"))
column_lengths = [None] * num_col
for i in range(num_col):
    column_lengths[i] = input("What is the length of column:" + str(i) + "?")

for filename in os.listdir(os.getcwd()):
    if filename.endswith(".xlsx"):  
        myfile = xlrd.open_workbook(filename)
        new_file = filename[:-5] + ".txt"
        new = open(new_file,"w")
        mysheet = myfile.sheet_by_index(0)
        for rownum in range(mysheet.nrows):
            for columnnum in range(len(column_lengths)):
                value = str(mysheet.cell_value(rownum, columnnum))
                value = add_spaces(value, column_lengths[columnnum])
                new.write(value)
            new.write("\n") 
    