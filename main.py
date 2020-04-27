import xlrd
import xlrd.sheet
import csv
import sys
import os
import configparser
import shutil

def add_spaces(value, num):
    while len(value) < int(num):
        value = value + " "
    return value

def no_split(column_lengths, retain_headers, retain_footers, ext_txt, custom_header, custom_footer):
    for filename in os.listdir(os.getcwd()):
        if filename.endswith(".xlsx"):  
            myfile = xlrd.open_workbook(filename)
            if ext_txt == "txt":
                new_file = "output/" + filename[:-5] + ".txt"
            else:
                new_file = "output/" + filename[:-5] + ".ext"
            new = open(new_file,"w")
            header = ""
            footer = ""
            mysheet = myfile.sheet_by_index(0)
            if (retain_headers == "yes"):
                header = str(mysheet.cell_value(0, 0))
                new.write(header)
                new.write("\n")
            if (custom_header != "no"):
                new.write(custom_header)
                new.write("\n")
            for rownum in range(mysheet.nrows):
                for columnnum in range(len(column_lengths)):
                    value = str(mysheet.cell_value(rownum, columnnum))
                    value = add_spaces(value, column_lengths[columnnum])
                    new.write(value)
                    footer = str(mysheet.cell_value(rownum, 0))
                new.write("\n") 
            if (retain_footers == "yes"):
                for key in batches:
                    curr = batches.get(key)
                    curr.write(footer)
            if (custom_footer != "no"):
                for key in batches:
                    curr = batches.get(key)
                    curr.write(custom_footer)

def split_batch(column_lengths, num_split, retain_headers, retain_footers, ext_txt, custom_header, custom_footer):
    batches = {}
    for filename in os.listdir(os.getcwd()):
        if filename.endswith(".xlsx"):  
            myfile = xlrd.open_workbook(filename)
            if ext_txt == "txt":
                new_file = "output/" + filename[:-5] + ".txt"
            else:
                new_file = "output/" + filename[:-5] + ".ext"
            new = open(new_file,"w")
            mysheet = myfile.sheet_by_index(0)
            header = ""
            footer = ""
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
                    if ext_txt == "txt":
                        new_file = "output/" + str(batch_num) + ".txt"
                    else:
                        new_file = "output/" + str(batch_num) + ".ext"
                    batches[batch_num] = open(new_file,"w")
                    new = batches.get(batch_num)
                    if (retain_headers == "yes"):
                        new.write(header)
                        new.write("\n")
                    if (custom_header != "no"):
                        new.write(custom_header)
                        new.write("\n")
                    for columnnum in range(len(column_lengths)):
                        value = str(mysheet.cell_value(rownum, columnnum))
                        value = add_spaces(value, column_lengths[columnnum])
                        new.write(value)
                footer = str(mysheet.cell_value(rownum, 0))
                new.write("\n") 
            if (retain_footers == "yes"):
                for key in batches:
                    curr = batches.get(key)
                    curr.write(footer)
            if (custom_footer != "no"):
                for key in batches:
                    curr = batches.get(key)
                    curr.write(custom_footer)

path = os.getcwd() + "/output"
try:
    os.mkdir(path)
except OSError:
    shutil.rmtree(path)
    os.mkdir(path)
config = configparser.ConfigParser()
config.read("config.ini")
column_lengths = config['Main']['ColumnLengths'].split()
retain_headers = config['Main']['RetainHeader']
retain_footers = config['Main']['RetainFooter']
ext_txt = config['Main']['ExtOrTxt']
custom_header = config['Main']['CustomHeader']
custom_footer = config['Main']['CustomFooter']
#config_file = open("config.txt", "r")
#config_lines = config_file.read().splitlines()
#column_lengths = config_lines[0].split()
#retain_headers = config_lines[2]
#retain_footers = config_lines[3]
#ext_txt = config_lines[4]
for i in range(len(column_lengths)):
    column_lengths[i] = int(column_lengths[i])
num_split = int(config['Main']['SplitColumn'])
if (num_split == -1):
    no_split(column_lengths, retain_headers, retain_footers, ext_txt, custom_header, custom_footer)
else:
    split_batch(column_lengths, num_split, retain_headers, retain_footers, ext_txt, custom_header, custom_footer)