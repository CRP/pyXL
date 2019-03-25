#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
tools which allow to manipulate excel ranges encapsulating applescript/win32/xlsxwriter instructions with same interface

switch the excel engine, the same routine may use different engines without changes to the code

For each engine the module exposes the following objects:
- Excel: an abstract object encapsulating the Excel application session
- Workbook: an abstract object encapsulating an Excel file
- Sheet: an abstract object encapsulating a sheet in a Workbook
- Rng: an abstract object encapsulating a range on a Sheet
and functions:
- df2rng: write a pandas object to a range
- rng2arr: read the content of a range into a list
- rng2df: read the content of a range into a pandas object

See the docstrings for these objects for more details.


example:

x=XL.Excel() # this is just a reference to excel and allows to set a few settings, such as calculation etc.

wb=x.create_wb() # create a new workbook, returns instance of Workbook object

sh=wb.create_sheet('pippo') # create a new sheet named "pippo", returns instance of Sheet object

wb.sheets # return a list of sheets

sh2=wb.sheets[1] # get reference of sheet 1

r=sh.arng('B2') # access a range on the sheet

r #prints current "coordinates"

temp=TS.DataFrame(np.random.randn(30,4),columns=list('abcd'))
r.from_pandas(temp) #write data to current sheet

r #coordinates have changed!!

r.format_range({'b':'@','d':'0.0000'},{'c':40}) #do some formatting

r.sort('b') # do some sorting

r.to_pandas() #read data from current sheet

"""

excelpath='/Applications/Microsoft Excel.app'

from pyXL.excel import *
switch_engine('applescript')
