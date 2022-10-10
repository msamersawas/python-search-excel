python-search-excel
Python script to search for text inside all Excel files in a directory. Python library openxl is used.

Usage
Search is case insensitive.

To search all Excel files under c:\data for text "qtr 1" , from a command prompt type:

python.exe find_text_in_workbooks_folder.py c:\data "qtr1"

To search all Excel files in current directory "." for text "qtr 1", from a command prompt type:

python.exe find_text_in_workbooks_folder.py . "qtr1"
