#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlsxwriter as XLW
"""
:mod:`excel_mac_as` -- Excel-Applescript wrapper
================================================

..module:: excel_mac_as
:platform: Mac
:synopsis: a class encapsulating access to Excel range object via Applescript
..moduleauthor:: Christian Prinoth < c.prinoth@quaestiocapital.com >

TODO:
* investigate applescript colorscale bug
* range intersection
* range union
* create excel charts

"""
import pandas as _pd
import os as _os
import numpy as _np
from copy import copy as _copy

from pyXL.excel_utils import _cr2a, _a2cr, _n2x, _x2n, _splitaddr, _df2outline, _isrow, _iscol, _isnumeric,_df_to_ll

class Rng:
    """
    class which allows to manipulate excel ranges encapsulating applescript instructions

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

    """

    def __repr__(self):
        """
        print out coordinates of the range object
        :return:
        """
        return "Workbook: %s\tSheet: %s\tRange: %s" % (self.sheet.workbook.name, self.sheet.name, self.address)

    def __init__(self, address=None, sheet=None, col=None, row=None):
        """
        Initialize a range object from a given set of address, sheet name and workbook name
        if any of these is not given, then they are fetched using "active" command
        :param address: range address, using A1 style (RC is not yet supported)
        :param sheet: name of the sheet
        :param workbook: name of the workbook
        """
        self.sheet = None
        self.address = None

        if address is not None or col is not None or row is not None:
            if address is not None:
                self.address = address.replace('$', '').replace('"', '')
            elif row is None and col is not None:
                c = _n2x(col)
                self.address = '%s:%s' % (c, c)
            elif row is not None and col is None:
                self.address = '%i:%i' % (row, row)
            else:
                self.address = '%s%i' % (_n2x(col), row)
            if sheet is not None:
                self.sheet = sheet
        self._set_address()

    def _set_address(self):
        if self.sheet is None:
            self.sheet = Sheet()
        if self.address is None:
            raise Exception("non dovremmo mai finire qui")

    def arng(self, address=None, row=None, col=None):
        """
        method for quick access to different range on same sheet
        :param address:
        :param row: 1-based
        :param col: 1-based
        :return:
        """
        return Rng(address=address, sheet=self.sheet, row=row, col=col)

    def offset(self, r=0, c=0):
        """
        return new range object offset from the original by r rows and c columns
        :param r: number of rows to offset by
        :param c: number of columns to offset by
        :return: new range object
        """
        coords = _a2cr(self.address)
        if len(coords) == 2:
            newaddr = _cr2a(coords[0] + c, coords[1] + r)
        else:
            newaddr = _cr2a(coords[0] + c, coords[1] + r, coords[2] + c, coords[3] + r)
        return Rng(address=newaddr, sheet=self.sheet)

    def iloc(self, r=0, c=0):
        """
        return a cell in the range based on coordinates starting from left top cell
        :param r: row index
        :param c: columns index
        :return:
        """
        coords = _a2cr(self.address)
        newaddr = _cr2a(coords[0] + c, coords[1] + r)
        return Rng(address=newaddr, sheet=self.sheet)

    def resize(self, r=0, c=0, abs=True):
        """
        new range object with address with same top left coordinate but different size (see abs param)
        :param r:
        :param c:
        :param abs: if true, then r and c determine the new size, otherwise they are added to current size
        :return: new range object
        """
        coords = _a2cr(self.address)
        if len(coords) == 2: coords = coords + coords
        if abs:
            newaddr = _cr2a(coords[0], coords[1], coords[0] + max(0, c - 1), coords[1] + max(0, r - 1))
        else:
            newaddr = _cr2a(coords[0], coords[1], max(coords[0], coords[2] + c), max(coords[1], coords[3] + r))
        return Rng(address=newaddr, sheet=self.sheet)

    def row(self, idx):
        """
        range with given row of current range
        :param idx: indexing is 1-based, negative indices start from last row
        :return: new range object
        """
        coords = _a2cr(self.address)
        if len(coords) == 2:
            return _copy(self)
        else:
            newcoords = _copy(coords)
            if idx < 0:
                newcoords[1] = newcoords[3] + idx + 1
            else:
                newcoords[1] += idx - 1
            newcoords[3] = newcoords[1]
            newaddr = _cr2a(*newcoords)
            return Rng(address=newaddr, sheet=self.sheet)

    def column(self, idx):
        """
        range with given col of current range
        :param idx: indexing is 1-based, negative indices start from last col
        :return: new range object
        """
        coords = _a2cr(self.address)
        if len(coords) == 2:
            return _copy(self)
        else:
            newcoords = _copy(coords)
            if idx < 0:
                newcoords[0] = newcoords[2] + idx + 1
            else:
                newcoords[0] += idx - 1
            newcoords[2] = newcoords[0]
            newaddr = _cr2a(*newcoords)
            return Rng(address=newaddr, sheet=self.sheet)

    def format(self, fmt=None, halignment=None, valignment=None, wrap_text=False, **kwargs):
        """
        formats a range
        if fmt is None, return current format
        otherwise, you can also set alignment and wrap text
        :param fmt: excel format string, look at U.FX for examples
        :param halignment: right, left, center
        :param valignment: top, middle, bottom
        :param wrap_text: true or false
        :param kwargs: see following list
                Font    	Font type	        'font_name'
                            Font size	        'font_size'
                            Font color	        'font_color'
                            Bold	            'bold'
                            Italic	            'italic'
                            Underline	        'underline'
                            Strikeout	        'font_strikeout'
                            Super/Subscript	    'font_script'
                Number	    Numeric format	    'num_format'
                Protection	Lock cells	        'locked'
                            Hide formulas	    'hidden'
                Alignment	Horizontal align	'align'
                            Vertical align	    'valign'
                            Rotation	        'rotation'
                            Text wrap	        'text_wrap'
                            Justify last	    'text_justlast'
                            Center across	    'center_across'
                            Indentation	        'indent'
                            Shrink to fit	    'shrink'
                Pattern	Cell pattern	        'pattern'
                            Background color	'bg_color'
                            Foreground color	'fg_color'
                Border	    Cell border	        'border'
                            Bottom border	    'bottom'
                            Top border	        'top'
                            Left border	        'left'
                            Right border	    'right'
                            Border color	    'border_color'
                            Bottom color	    'bottom_color'
                            Top color	        'top_color'
                            Left color	        'left_color'
                            Right color	        'right_color'
        :return:
        """
        if self.address not in self.sheet.cell_formats.keys():
            self.sheet.cell_formats[self.address] = {}
        if fmt is not None:
            self.sheet.cell_formats[self.address]['num_format']=fmt
        if halignment is not None:
            self.sheet.cell_formats[self.address]['align']=halignment
        if valignment is not None:
            self.sheet.cell_formats[self.address]['valign']=valignment
        if wrap_text is not None:
            self.sheet.cell_formats[self.address]['text_wrap']=wrap_text
        for k,v in kwargs.items():
            self.sheet.cell_formats[self.address][k] = v

    def font_format(self, bold=False, italic=False, name='Calibri', size=12, color=(0, 0, 0)):
        """
        set properties of range fonts
        :param bold: true/false
        :param italic: true/false
        :param name: a font name, such as Calibri or Courier
        :param size: number
        :param color: RGB triplet
        :return: a list of current font properties
        """
        if self.address not in self.sheet.cell_formats.keys():
            self.sheet.cell_formats[self.address] = {}
        if bold is not None:
            self.sheet.cell_formats[self.address]['bold']=bold
        if italic is not None:
            self.sheet.cell_formats[self.address]['italic']=italic
        if name is not None:
            self.sheet.cell_formats[self.address]['font_name']=name
        if size is not None:
            self.sheet.cell_formats[self.address]['font_size']=size
        if color is not None:
            self.sheet.cell_formats[self.address]['font_color']=_rgb2xlcol(color)


    def filldown(self):
        """
        TODO
        fill down content of first row to rest of selection
        :return:
        """
        pass

    def color(self, col=None):
        """
        colors a range interior
        :param col: RGB triplet, or None to remove coloring
        :return:
        """
        if self.address not in self.sheet.cell_formats.keys():
            self.sheet.cell_formats[self.address] = {}
        if col is not None:
            self.sheet.cell_formats[self.address]['bg_color']=_rgb2xlcol(col)

    def value(self, v=None):
        """
        get or set the value of a range
        :param v: value to be set, if None current value is returned
        :return:
        """
        if v is not None:
            if isinstance(v, (_pd.DataFrame, _pd.Series)):
                return self.from_pandas(v)
            elif isinstance(v, (list, tuple, _np.ndarray)):
                temp = _pd.DataFrame(v)
                return self.from_pandas(temp, header=False, index=False)
            self.sheet.cell_data[self.address]=v
        else:
            pass

    def get_array(self, string_value=False):
        """
        get an excel range as a list of lists

        currently has problems with strings containing the { or } characters
        :return: list
        """
        pass

    def get_df(self, index=0, header=0):
        """
        return a range as dataframe
        :param index: None for no index, otherwise an integer specifying n columns to use as index
        :param header: None to avoid using columns, any other value to use first row as column header
        :return:
        """
        pass

    # def cell(self, value=None, formula=None, format=None, asarray=False):
    #     """
    #     TODO
    #     get or set value, formula and format of a cell
    #
    #     :param value:
    #     :param formula:
    #     :param format:
    #     :param asarray: specify if formula should be of array type
    #     :return:
    #     """
    #     pass

    def formula(self, f=None, asarray=False):
        """
        get or set the value of a range
        :param f: formula to be set, if None current formula is returned
        :param asarray: set array formula
        :return:
        """
        if f is not None:
            self.sheet.cell_data[self.address]='{'+f+'}' if asarray else f
        else:
            pass

    def get_selection(self):
        """
        TODO
        refresh range coordinates based on current selection
        this modifies the current instance of the object!
        :return:
        """
        pass

    def get_cells(self):
        """
        TODO
        return a list of all addresses of cells in range
        :return:
        """
        pass

    def from_pandas(self, pdobj, header=True, index=True, index_label=None, outline_string=None):
        """
                write a pandas object to excel via clipboard
        :param pdobj: any DataFrame or Series object

        see DataFrame.to_clipboard? for info on params below

        :param header: if False, strip header
        :param index: if False, strip index
        :param index_label: index header
        :param outline_string: a string used to identify outline main levels (eg " All")
        :return:
        """

        temp=_df_to_ll(pdobj,header=header,index=index)
        trange = self.resize(len(temp), len(temp[0]))
        coords=_a2cr(trange.address)
        for j,c in enumerate(range(coords[0],coords[2]+1)):
            for i,r in enumerate(range(coords[1], coords[3] + 1)):
                addr=_cr2a(c,r)
                self.sheet.cell_data[addr] = temp[i][j]

        self.address = trange.address
        if outline_string is not None:
            boundaries = _df2outline(pdobj, outline_string)
            self.outline(boundaries)
        return trange

    def to_pandas(self, index=1, header=1):
        """
        read excel data into a pandas object via clipboard

        WARNING: data will be read from the excel sheet "as-is", so format them correctly beforehand!!!!

        :param index_col: if None, do not reat index, else use first index_col columns
        :param header: if None, do not read header, esel use first row
        :param parse_dates: parse date columns
        :return: a DataFrame object
        """
        pass

    def clear_formats(self):
        """
        clear all formatting from range
        :return:
        """
        pass

    def delete(self, r=None, c=None):
        """
        delete entire rows or columns
        :param r: a (list of) row(s)
        :param c: a (list of) column(s)
        :return:
        """
        pass

    def insert(self, r=None, c=None):
        """
        insert rows or columns
        :param r: a (list of) row(s)
        :param c: a (list of) column(s)
        :return:
        """
        pass

    def clear_values(self):
        """
        clear all values from range
        :return:
        """
        pass

    def column_width(self, w):
        """
        change width of the column(s) of range
        :param w: width
        :return:
        """
        c1,r1,c2,r2=_a2cr(self.address)
        self.parent.ws.set_column(c1-1,c2-1,width=w)
        #self.sheet.workbook.parent.set_column(c1-1,c2-1,width=w)

    def row_height(self, h):
        """
        change height of the row(s) of range
        :param h: height
        :return:
        """
        c1,r1,c2,r2=_a2cr(self.address)
        self.parent.ws.set_column(r1-1,height=h)

    def curr_region(self):
        """
        get range of the current region
        :return: new range object
        """
        curr=_get_contiguous(self.address,self.sheet.cell_data.keys())
        return Rng(address=curr, sheet=self.sheet)

    def replace(self, val, repl_with, whole=False):
        """
        within the range, replace val with repl_with
        :param val: value to be looked for
        :param repl_with: value to replace with
        :return:
        """
        pass

    def sort(self, key1, order1=None, key2=None, order2=None, key3=None, order3=None, header=True):
        """
        sort data in a range, for now works only if header==True
        keys must be column header labels
        :param key1: header string
        :param order1: d/a
        :param key2:
        :param order2:
        :param key3:
        :param order3:
        :param header:
        :return:
        """
        pass

    def format_range(self, fmt_dict={}, cw_dict={}, columns=True):
        """
        TODO
        formats multiple columns (or rows) at once
        :param fmt_dict: dictionary where keys are column headers of the range (i.e. strings in the first row) while
                         values are excel formatting codes
                         fmt_dict keys may also be regular expressions which are then matched against column names
                         (eg .* may be used as wildcard key, i.e. any columns are matched against this)
        :param cw_dict: same fmt_dict but instead of format strings must contain column widths (row heights)
        :param columns: if True iterate over columns, else over rows
        :return:
        """
        pass


    def freeze_panes(self):
        """
        freezes panes at upper left cell of range
        :return:
        """
        c,r=_a2cr(self.address)
        self.parent.ws.freeze_panes(r-1,c-1)

    def color_scale(self, vmin=5, vmed=50, vmax=95, cv=5):
        """
        TODO
        apply 3 color scale formatting
        NOTE: THIS ONLY WORKS IF CP.XLAM ADDIN IS LOADED IN EXCEL!!!! as applescript handling of advanced
            conditional formatting is apparently broken

        :param vmin: minimum value
        :param vmed: median value
        :param vmax: maximum value
        :param cv: a number that determines what "value" in the previous params is, from the following list
            5 Percentile is used. (default)
            7 The longest data bar is proportional to the maximum value in the range.
            6 The shortest data bar is proportional to the minimum value in the range.
            4 Formula is used.
            2 Highest value from the list of values.
            1 Lowest value from the list of values.
            -1 No conditional value.
            0 Number is used.
            3 Percentage is used.
        :return:
        """
        pass

    def select(self):
        """
        select the range
        :return:
        """
        pass

    def highlight(self, condition='==', threshold=0.0, interiorcolor=(255, 0, 0)):
        """
        TODO
        highlight cells satisfying given condition
        :param condition: one of ==,!=,>,<,>=,<=
        :param threshold: a number
        :param interiorcolor: an RGB triple specifying the color, eg [255,0,0] is red
        :return:
        """
        pass

    def col_dict(self):
        """
        TODO
        given a range with a header row,
        return a dictionary of range objects, each representing a column of the current range
        :return: dict, where keys are header strings, while values are column range objects
        """
        pass

    def autofit_rows(self):
        """
        autofit row height
        :return:
        """
        pass

    def autofit_cols(self):
        """
        autofit column width
        :return:
        """
        pass

    def entire_row(self):
        """
        get entire row(s) of current range
        :return: new object
        """
        c = _a2cr(self.address)
        if len(c) == 2: c += c
        cc = '%s:%s' % (c[1], c[3])
        return Rng(address=cc, sheet=self.sheet)

    def entire_col(self):
        """
        get entire row(s) of current range
        :return: new object
        """
        c = _a2cr(self.address)
        if len(c) == 2: c += c
        cc = '%s:%s' % (_n2x(c[0]), _n2x(c[2]))
        return Rng(address=cc, sheet=self.sheet)

    def activate(self):
        """
        activate range
        :return:
        """
        pass

    def propagate_format(self, col=True):
        """
        propagate formatting of first column (row) to subsequent columns (rows)
        :param col: True for columns, False for rows
        :return:
        """
        pass

    def outline(self, boundaries):
        """
        TODO
        group rows as defined by boundaries object
        :param boundaries: dictionary, where keys are group "main level" and values is a list of two
                           identifying subrows referring to main level
        :return:
        """
        import six
        top=2**20
        lvl=0
        for k,[f, l] in six.iteritems(boundaries):
            if l<top:
                top=l # this only works if boundaries are correctly sorted
                lvl+=1
            r=self.offset(r=k).resize(r=l-k).entire_row()
            coords = _a2cr(r.address)
            for row in range(coords[1], coords[3] + 1): # if I don't handle each row separately levels get overwritten!
                addr='%i:%i'%(row,row)
                if addr not in self.sheet.cell_options.keys(): self.sheet.cell_options[addr] = {}
                if 'level' in self.sheet.cell_options[addr].keys():
                    lvl=max(lvl,self.sheet.cell_options[addr]['level'])
                self.sheet.cell_options[addr]['level']=lvl
                if lvl>2:
                    self.sheet.cell_options[addr]['collapsed']=True
                elif lvl==2:
                    self.sheet.cell_options[addr]['hidden'] = True
            r=self.offset(r=k-1).row(1)
            r.font_format(bold=True)

    def show_levels(self, n=2):
        """
        TODO
        set level of outline to show
        :param n:
        :return:
        """
        pass

    def goal_seek(self, target, r_mod):
        """
        set value of range of self to target by changing range r_mod
        :param target: the target value for current range
        :param r_mod: the range to modify, or an integer with the column offset from the current cell
        :return:
        """
        pass

    def paste_fig(self, figpath, w, h):
        """
        paste a figure from a file onto an excel sheet, setting width and height as specified
        location will be the top left corner of the current range

        :param figpath: posix path of file
        :param w: width in pixels
        :param h: height in pixels
        :return:
        """
        self.sheet.images[self.address]=figpath

    def subrng(self, t, l, nr=1, nc=1):
        """
        given a range returns a subrange defined by relative coordinates
        :param t: row offset from current top row
        :param l: column offset from current top column
        :param nr: number of rows in subrange
        :param nc: number of columns in subrange
        :return: range object
        """
        coords = _a2cr(self.address)
        newaddr = _cr2a(coords[0] + l, coords[1] + t, coords[0] + l + nc-1, coords[1] + t + nr-1)
        return Rng(address=newaddr, sheet=self.sheet)

    def subtotal(self, groupby, totals, aggfunc='sum'):
        """
        :param groupby:
        :param totals:
        :param aggfunc:
        :return:
        """
        pass

    def size(self):
        """
        return size of range
        :return: columns, rows
        """
        temp=_a2cr(self.address)
        return temp[2]-temp[0]+1,temp[3]-temp[1]+1


class Excel():
    """
    basic wrapper of Excel application, providing some methods to perform simple automation, such as
    creating/opening/closing workbooks
    """

    def __repr__(self):
        """
        print out coordinates of the range object
        :return:
        """
        return 'Excel application, currently %i workbooks are open' % len(self.workbooks)

    def __init__(self):
        self.awb=None
        self.workbooks = []
        self._calculation_manual = False
        self.refresh_workbook_list()

    def refresh_workbook_list(self):
        """
        make sure that object is consistent with current state of excel
        this needs to be called if, during an interactive session, users create/delete workbooks manually,
        as the Excel object has no way to know what the user does
        :return:
        """
        pass

    def active_workbook(self):
        """
        return the active workbook
        :return:
        """
        raise Exception('xlsxwriter does not support editing existing files')
        #return self.workbooks[self.awb]

    def active_range(self):
        """
        return the active range
        :return:
        """
        wb = self.workbooks(self.awb)
        ws=wb.get_active_sheet()
        return ws['A1']

    def create_wb(self, name='Workboox.xlsx'):
        """
        create a new workbook
        :return:
        """
        return Workbook(parent=self, name=name)

    def get_wb(self, name):
        """
        get a reference to a workbook based on its name
        :param name:
        :return:
        """
        found=False
        for wb in self.workbooks:
            if wb.name==name:
                found=True
                break
        if found:
            return wb
        else:
            raise Exception("No workbook named %s"%name)

    def calculation(self, manual=True):
        """
        set calculation and screenupdating of excel to manual or automatic
        :param manual:
        :return:
        """
        self.wb.set_calc_mode('manual' if manual else 'auto')

    def open_wb(self, fpath):
        """
        open a workbook given its path
        :param fpath:
        :return:
        """
        raise Exception("xlsxwriter does not allow working with existing files")

class Workbook():
    """
    an object representing an Excel workbook, and providing a few methods to automate it
    """

    def __repr__(self):
        """
        print out coordinates of the range object
        :return:
        """
        return "Workbook object '%s', has %i sheets" % (self.name, len(self.sheets))

    def __init__(self, existing=None, parent=None, name='Workbook.xlsx'):

        self.name = None
        self.wb = None # reference to actual xlsxwriter object
        self.parent = None
        self.sheets = []
        self.path=name

        if parent is not None:
            self.parent = parent
        elif self.parent is None:
            self.parent = Excel()

        if existing is None:
            name=name
            #self.wb=XLW.Workbook(name,{'nan_inf_to_errors': True,'default_date_format': 'yyyy-mm-dd',})
            self.name = name
        else:
            #self.wb = existing[0]
            self.name=existing[1]
        self.parent.workbooks.append(self)
        #make sure that workbook has always at least one sheet
        if len(self.sheets)==0: self.create_sheet('Sheet')

    def create_sheet(self, name='Sheet'):
        """
        create a new sheet in the current workbook
        :param name:
        :return:
        """
        return Sheet(name=name, workbook=self)

    def refresh_sheet_list(self):
        """
        make sure that object is consistent with current state of excel
        this needs to be called if, during an interactive session, users create/delete sheets manually,
        as the Excel object has no way to know what the user does
        :return:
        """
        pass
        #temp=self.wb.worksheets
        #for t in temp:
        #    Sheet(existing=t, workbook=self)

    def saveas(self, fpath):
        """
        save a workbook into a different file (silently overwrites existing file with same name!!!)
        :param fpath:
        :return:
        """
        #self.wb.close()
        self.name = _os.path.basename(fpath)
        self.path=fpath

    def save(self, fpath):
        """
        save workbook
        :param fpath:
        :return:
        """
        #print("filename can not be changed, use close method to save")
        #self.wb.close()
        self.name = _os.path.basename(fpath)
        self.path=fpath

    def close(self):
        """
        close a workbook without saving it
        :return:
        """
        #create workbook
        self.wb=XLW.Workbook(self.path,{'nan_inf_to_errors': True,'default_date_format': 'yyyy-mm-dd',})
        #create sheets
        for sheet in self.sheets:
            sheet.ws=self.wb.add_worksheet(sheet.name)
            sheet.ws.outline_settings(outline_below=False,outline_right=False)
        # write all data to all sheets
        for sheet in self.sheets:
            cells=list(set(sheet.cell_data.keys()).
                       union(set(sheet.cell_formats.keys())).
                       union(set(sheet.cell_options.keys()))
                       )
            rows=list(filter(lambda x: _isrow(x),cells))
            for row in rows:
                if row in sheet.cell_formats.keys():
                    format = sheet.cell_formats[row]
                    format = self.wb.add_format(format)
                else:
                    format=None
                if row in sheet.cell_options.keys():
                    options= sheet.cell_options[row]
                else:
                    options={}
                coords=_a2cr(row)
                for r in range(coords[1],coords[3]+1):
                    sheet.ws.set_row(r,None,cell_format=format,options=options)
            cols=list(filter(lambda x: _iscol(x),cells))
            for col in cols:
                if col in sheet.cell_formats.keys():
                    format = sheet.cell_formats[col]
                    format = self.wb.add_format(format)
                else:
                    format=None
                if col in sheet.cell_options.keys():
                    options= sheet.cell_options[row]
                else:
                    options={}
                coords=_a2cr(row)
                for c in range(coords[0],coords[2]+1):
                    sheet.ws.set_col(c,None,cell_format=format,options=options)
            cells=list(set(cells)-set(rows)-set(cols))
            for cell in cells:
                if cell in sheet.cell_formats.keys():
                    format=sheet.cell_formats[cell]
                    format = self.wb.add_format(format)
                else:
                    format=None
                if cell in sheet.cell_data.keys():
                    value=sheet.cell_data[cell]

                    if len(str(value))>0:

                        if str(value)[0]=='=':
                            sheet.ws.write_formula(cell, value, format)
                        elif str(value)[0]== '{':
                            sheet.ws.write_array_formula(cell, value, format)
                        else:
                            coords=_a2cr(cell)
                            if len(coords)==2:
                                coords=[coords[0],coords[1],coords[0],coords[1]]
                            for c in range(coords[0],coords[2]+1):
                                for r in range(coords[1],coords[3]+1):
                                    sheet.ws.write(r - 1, c - 1, value, format)

                    else:
                        coords=_a2cr(cell)
                        if len(coords)==2:
                            coords=[coords[0],coords[1],coords[0],coords[1]]
                        for c in range(coords[0],coords[2]+1):
                            for r in range(coords[1],coords[3]+1):
                                sheet.ws.write(r - 1, c - 1, value, format)


                else:
                    sheet.ws.write_blank(cell, None, format)
            for addr, figpath in sheet.images.items():
                c, r = _a2cr(addr)
                sheet.ws.insert_image(r, c, figpath)

        self.parent.workbooks.remove(self)
        self.wb.close()

    def get_sheet(self, name):
        """
        get a reference to a sheet object given a name
        :param name:
        :return:
        """
        for i, sh in enumerate(self.sheets):
            if sh.name == name:
                break
        if len(self.sheets) == 0 or i > len(self.sheets):
            raise Exception("there is no sheet %s" % name)
        else:
            return sh

class Sheet():
    """
    an object representing an Excel sheet, and providing a few methods to automate it
    """

    def __repr__(self):
        """
        :return:
        """
        return "Worksheet object %s, owned by '%s'" % (self.name, self.workbook.name)

    def __init__(self, existing=None, workbook=None, name='Sheet'):

        self.workbook = None
        self.ws = None #reference to the actual xlsxwriter worksheet object
        self.name = None
        self.rng = None
        self.cell_data = {}
        self.cell_formats = {}
        self.cell_options = {}
        self.images = {}

        if workbook is None:
            self.workbook = Workbook(name='WB_'+name)
            existing = self.workbook.sheets[0].name
        else:
            self.workbook = workbook

        if existing is None:
            uname=name
            slist = [ws.name for ws in self.workbook.sheets]
            i = 0
            while uname in slist:
                i += 1
                uname = name + '(%i)' % i
            self.name = uname
        else:
            self.name = existing
        self.rng = Rng('A1', sheet=self)
        self.workbook.sheets.append(self)

    def arng(self, address=None, row=None, col=None):
        """
        access a range on the sheet, providing either address in A1 format, or a row and/or a column
        :param address: string in A1 format
        :param row: integer, 1-based
        :param col: integer, 1-based
        :return:
        """
        self.rng = Rng(address=address, row=row, col=col, sheet=self)
        return self.rng

    def rename(self, name):
        """
        change the name of the current sheet
        :param name:
        :return:
        """
        pass

    def unprotect(self):
        """
        remove protection (only if no password!)
        :return:
        """
        pass

    def protect(self):
        """
        activate protection (without password!)
        :return:
        """
        pass

    def delete_shapes(self):
        """
        delete all shape objects on the current sheet
        :return:
        """
        pass

    # def get_values_formulas_formats(self, *rngs):
    #     """
    #     traverses a range and returns its contents as a dictionary
    #     keys are cell addresses, values are content, formulas and formats
    #
    #     :param rngs: one or more range addresses
    #     :return: 3 dicts
    #     """
    #     raise Exception()
    #     outvalues = {}
    #     outformulas = {}
    #     outformats = {}
    #     for r in rngs:
    #         rn = Rng(r, sheet=self)
    #         cells = rn.get_cells()
    #         for c in cells:
    #             print(c)
    #             v, s, f, t = rn.arng(c).cell()
    #             if v != '':
    #                 outvalues[c] = v
    #             if len(f) > 0:
    #                 outformulas[c] = f
    #             if len(t) > 0:
    #                 outformats[c] = t
    #     return outvalues, outformulas, outformats

    def set_values_formulas_formats(self, values_dict=None, formats_dict=None,
                                    formulas_dict=None, arrformulas_dict=None):

        """
        traverses a range and sets its contents from a dictionary
        keys are cell addresses, values are content, formulas and formats

        :param values_dict:
        :param formats_dict:
        :param formulas_dict:
        :param arrformulas_dict:
        :return:
        """

        if values_dict is not None:
            for addr, v in values_dict.items():
                self.cell_data[addr]=v
        if formats_dict is not None:
            for addr, v in formats_dict.items():
                if addr not in self.cell_formats.keys(): self.cell_formats[addr]={}
                self.cell_formats[addr] = v
        if formulas_dict is not None:
            for addr, v in formulas_dict.items():
                if addr not in self.cell_formats.keys(): self.cell_formats[addr]={}
                self.cell_formats[addr] = v
        if arrformulas_dict is not None:
            for addr, v in arrformulas_dict.items():
                if addr not in self.cell_formats.keys(): self.cell_formats[addr]={}
                self.cell_formats[addr] = '{'+v+'}'


def _rgb2xlcol(rgb):
    return '#%02x%02x%02x' % tuple([255,0,0])

def _get_contiguous(address, cells):

    #get current top left and bottom right
    boundaries=_a2cr(address,f4=True)
    #look for any cells contiguous to current rectangle
    for c in cells:
        cc=_a2cr(c,f4=True) #coords of current cell
        if _is_contiguous(boundaries,cc):
            # expand top left bottom right
            boundaries=_expand_range(boundaries,cc)
    return _cr2a(*boundaries)

def _is_contiguous(a,b):
    """
    return True if b is contiguous to a
    :param a:
    :param b:
    :return:
    """
    if len(a)==2: a=[a[0],a[1],a[0],a[1]]
    if len(b)==2: b=[b[0],b[1],b[0],b[1]]
    #coords are left, top, right, bottom
    x1=b[2]>=a[0]-1
    x2=b[0]<=a[2]+1
    x3=b[3]>=a[1]-1
    x4=b[1]<=a[3]+1
    return all([x1,x2,x3,x4])

def _expand_range(a,b):
    return [min(a[0],b[0]),min(a[1],b[1]),max(a[2],b[2]),max(a[3],b[3])]
