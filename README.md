# Excel to Notepad
Small program which splits an excel file into a space delimited file that is human-readable even in default notepad applications

This was originally written for my mother to assist her with her work but I figured it might be useful to someone else.

# Usage:
Clone the repo and put the excel files which you want to convert into the repo folder as the script and run the script using
`python main.py`
The files will be converted and placed in the output folder in the current working directory and will retain the same name (with the extension changed to .txt)

# Config:
There is an optional config mode in which the program will read from the config file instead of requiring the user to input each attribute. The config.ini is formatted as follows
```
[Main]
ColumnLengths = a b c ...
SplitColumn = s
RetainHeader = yes|no
RetainFooter = yes|no
ExtOrTxt = ext/txt
```
`ColumnLengths` is the length of each column you would like to conver
- `a`, `b` and `c` are space delimited column lengths of columns A, B, C, and D in excel respectively. 
- You can include an infinite number of columns. 

`SplitColumn` is the column on which to split the data into separate files
- `s` represents the column number on which to split (starting at 1)
- Leave `s` as -1 to not split the file. 

`RetainHeader` is for whether or not to retain the header of the excel file
- Specify `yes` to retain header and `no` to discard it

`RetainFooter` is for whether or not to retain the footer of the excel file
- Specify `yes` to retain footer and `no` to discard it

`ExtOrTxt` is for whether you'd like the output saved as a `.ext` or `.txt` file. 
- Specify `ext` or `txt` for `.ext` and `.txt` respectively
