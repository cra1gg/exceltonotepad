# Excel to Notepad
Small program which splits an excel file into a space delimited file that is human-readable even in default notepad applications

# Usage:
Clone the repo and put the excel files which you want to convert into the repo folder as the script and run the script using
`python main.py`
The files will be converted and placed in the same directory and will retain the same name (with the extension changed to .txt)

# Config:
There is an optional config mode in which the program will read from the config file instead of requiring the user to input each attribute. The config.txt is formatted as follows
```
a b c d ...
p
```
where x, y, z, and p are space delimited column lengths of columns A, B, C, and D in excel accordinly. You can include an infinite number of columns. p is the column on which to split the data into separate files, leave this as -1 to ignore.
