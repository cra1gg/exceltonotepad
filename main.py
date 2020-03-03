import xlrd
import xlrd.sheet
import csv
import sys
import os

def add_spaces(value, num):
    while len(value) < int(num):
        value = value + " "
    return value

def no_split(column_lengths, retain_headers):
    for filename in os.listdir(os.getcwd()):
        if filename.endswith(".xlsx"):  
            myfile = xlrd.open_workbook(filename)
            new_file = filename[:-5] + ".txt"
            new = open(new_file,"w")
            mysheet = myfile.sheet_by_index(0)
            if (retain_headers == "yes"):
                header = str(mysheet.cell_value(0, 0))
            new.write(header)
            new.write("\n")
            for rownum in range(mysheet.nrows):
                for columnnum in range(len(column_lengths)):
                    value = str(mysheet.cell_value(rownum, columnnum))
                    value = add_spaces(value, column_lengths[columnnum])
                    new.write(value)
                new.write("\n") 

def split_batch(column_lengths, num_split, retain_headers):
    batches = {}
    for filename in os.listdir(os.getcwd()):
        if filename.endswith(".xlsx"):  
            myfile = xlrd.open_workbook(filename)
            new_file = filename[:-5] + ".txt"
            new = open(new_file,"w")
            mysheet = myfile.sheet_by_index(0)
            header = ""
            if (retain_headers == "yes"):
                header = str(mysheet.cell_value(0, 0))
            for rownum in range(mysheet.nrows):
                batch_num = mysheet.cell_value(rownum, num_split)
                if batch_num in batches:
                    new = batches.get(batch_num)
                    for columnnum in range(len(column_lengths)):
                        value = str(mysheet.cell_value(rownum, columnnum))
                        value = add_spaces(value, column_lengths[columnnum])
                        new.write(value)
                else:
                    new_file = filename[:-5] + "-BATCH " + str(batch_num) + ".txt"
                    batches[batch_num] = open(new_file,"w")
                    new = batches.get(batch_num)
                    new.write(header)
                    new.write("\n")
                    for columnnum in range(len(column_lengths)):
                        value = str(mysheet.cell_value(rownum, columnnum))
                        value = add_spaces(value, column_lengths[columnnum])
                        new.write(value)
                new.write("\n") 

config = input("Config mode or input mode? Enter c for config or i for input? ")
if config == "i":
    num_col = int(input("How many columns are there?: "))
    column_lengths = [None] * num_col
    for i in range(num_col):
        column_lengths[i] = input("What is the length of column " + str(i) + "?: ")

    num_split = int(input("Does this need to be split into batches? If so, please enter the column number on which to split (Enter -1 for n/a): "))
    retain_headers = input("Does this need to retain  headers and footers? (yes/no)")
    if (num_split == -1):
        no_split(column_lengths, retain_headers)
    else:
        split_batch(column_lengths, num_split, retain_headers)
else:
    config_file = open("config.txt", "r")
    config_lines = config_file.read().splitlines()
    column_lengths = config_lines[0].split()
    retain_headers = config_lines[2]
    for i in range(len(column_lengths)):
        column_lengths[i] = int(column_lengths[i])
    num_split = int(config_lines[1])
    if (num_split == -1):
        no_split(column_lengths, retain_headers)
    else:
        split_batch(column_lengths, num_split, retain_headers)