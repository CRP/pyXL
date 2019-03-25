#!/usr/bin/env python
# -*- coding: utf-8 -*-

try:
    import win32com.client as _win
    import pywintypes as _pywintypes
except:
    raise Exception()

import pandas as _pd
import os as _os
import datetime as _datetime
import pytz as _pytz
import numpy as _np
from copy import copy as _copy

from pyXL.excel_utils import _cr2a, _a2cr, _n2x, _x2n, _splitaddr, _df2outline, _isnumeric,_df_to_ll


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

    r.to_pandas() #read data from current sheet

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
        if any of these is not given, then they are fetched using "active" logic
        :param address: range address, using A1 style (RC is not yet supported)
        :param sheet: name of the sheet
        :param workbook: name of the workbook
        """
        self.range=None
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
            addr=self.sheet.ws.Selection.GetAddress(0, 0)
            if self.address is None:
                self.address = addr
            if self.sheet is None:
                self.sheet = Sheet(existing=self.sheet.ws)
        self.range=self.sheet.ws.Range(self.address)

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

    def format(self, fmt=None, halignment=None, valignment=None, wrap_text=False):
        """
        formats a range
        if fmt is None, return current format
        otherwise, you can also set alignment and wrap text
        :param fmt: excel format string, look at U.FX for examples
        :param halignment: right, left, center
        :param valignment: top, middle, bottom
        :param wrap_text: true or false
        :return:
        """
        if fmt is not None:
            self.range.NumberFormat=fmt
        elif halignment is not None:
            pass
        elif valignment is not None:
            pass
        elif wrap_text:
            pass
        else:
            return self.range.NumberFormat

    def filldown(self):
        """
        fill down content of first row to rest of selection
        :return:
        """
        self.range.FillDown()

    def color(self, col=None):
        """
        colors a range interior
        :param col: RGB triplet, or None to remove coloring
        :return:
        """
        self.range.Interior.Color = _rgb2xlcol(col)

    def value(self, v=None):
        """
        get or set the value of a range
        :param v: value to be set, if None current value is returned
        :return:
        """
        if v is None:
            out = self.range.Value
            out=_parse_windates(out)
            return out
        else:
            if _isnumeric(v):
                self.range.Value = v
            elif isinstance(v, str):
                self.range.Value = v
            elif isinstance(v, (_datetime.date, _datetime.datetime)):
                self.range.Value = _dt2pywintime(v)
            elif isinstance(v, (_pd.DataFrame, _pd.Series)):
                return self.from_pandas(v)
            elif isinstance(v, (list, tuple, _np.ndarray)):
                temp = _pd.DataFrame(v)
                return self.from_pandas(temp, header=False, index=False)
            else:
                raise Exception('Unhandled datatype')

    def get_array(self, string_value=False):
        """
        get an excel range as a list of lists

        :return: list
        """
        val=self.value()
        val=_parse_windates(val)
        return val

    def to_pandas(self, index=1, header=1):
        """
        return a range as dataframe
        :param index: None for no index, otherwise an integer specifying first n columns to use as index
        :param header: None to avoid using columns, any other value to use first n rows as column header
        :return:
        """
        temp=self.get_array()
        if header is None:
            temp = _pd.DataFrame(_pd.np.array(temp))
        elif header==1:
            hdr = temp[0]
            temp = _pd.DataFrame(_pd.np.array(temp[header:]), columns=hdr)
        elif header>1:
            hdr=_pd.MultiIndex.from_tuples(temp[:header])
            temp = _pd.DataFrame(_pd.np.array(temp[header:]), columns=hdr)
        else: raise Exception()

        if index is not None:
            temp = temp.set_index(temp.columns.tolist()[:index])
        return temp

    # def cell(self, value=None, formula=None, format=None, asarray=False):
    #     """
    #     get or set value, formula and format of a cell
    #
    #     :param value:
    #     :param formula:
    #     :param format:
    #     :param asarray: specify if formula should be of array type
    #     :return:
    #     """
    #     if value is not None:
    #         self.value(value)
    #     if formula is not None:
    #         self.formula(formula,asarray=asarray)
    #     if format is not None:
    #         self.format(format)
    #
    #     if value is None and formula is None and format is None:
    #         return self.value(), self.range.StringValue, self.formula(), self.format()

    def formula(self, f=None, asarray=False):
        """
        get or set the value of a range
        :param f: formula to be set, if None current formula is returned
        :param asarray: set array formula
        :return:
        """
        if f is not None:
            if asarray:
                self.range.FormulaArray=f
            else:
                self.range.Formula=f
        else:
            return self.range.Formula

    def get_selection(self):
        """
        refresh range coordinates based on current selection
        this modifies the current instance of the object!
        :return:
        """
        self.address = self.sheet.workbook.parent.app.Selection.GetAddress(0,0)
        self._set_address()

    def get_cells(self):
        """
        TODO
        return a list of all addresses of cells in range
        :return:
        """
        temp = self.range.Cells()
        return [x.GetAddress(0,0) for x in temp]

    def from_pandas(self, pdobj, header=True, index=True, index_label=None, outline_string=None):
        """
                write a pandas object to excel
        :param pdobj: any DataFrame or Series object

        :param header: if False, strip header
        :param index: if False, strip index
        :param index_label: index header
        :param outline_string: a string used to identify outline main levels (eg " All")
        :return:
        """
        temp = _df_to_ll(pdobj,header=header, index=index)
        temp = _fix_4_win(temp)
        trange = self.resize(len(temp), len(temp[0]))
        trange.range.Value=temp
        self.address = trange.address
        self.range = self.sheet.ws.Range(self.address)
        if outline_string is not None:
            boundaries = _df2outline(pdobj, outline_string)
            self.outline(boundaries)
        return temp

    def clear_formats(self):
        """
        clear all formatting from range
        :return:
        """
        self.range.ClearFormats()

    def delete(self, r=None, c=None):
        """
        delete entire rows or columns
        :param r: a (list of) row(s)
        :param c: a (list of) column(s)
        :return:
        """
        assert (r is None) ^ (c is None), "Either r or c must be specified, not both!"

        if r is not None:
            if not isinstance(r, (tuple, list)): r = [r]
            for rr in r:
                self.iloc(r=rr).EntireRow.Delete(Shift=_pywintypes.xlUp)
        if c is not None:
            if not isinstance(c, (tuple, list)): c = [c]
            for rr in c:
                self.iloc(c=rr).EntireCol.Delete(Shift=_pywintypes.xlLeft)


    def insert(self, r=None, c=None):
        """
        insert rows or columns
        :param r: a (list of) row(s)
        :param c: a (list of) column(s)
        :return:
        """
        return

        assert (r is None) ^ (c is None), "Either r or c must be specified, not both!"
        if r is not None:
            if not isinstance(r, (tuple, list)): r = [r]
            for rr in r:
                self.iloc(r=rr).EntireRow.Insert(Shift=_pywintypes.xlDown)
        if c is not None:
            if not isinstance(c, (tuple, list)): c = [c]
            for rr in c:
                self.iloc(c=rr).EntireCol.Insert(Shift=_pywintypes.xlRight)

    def clear_values(self):
        """
        clear all values from range
        :return:
        """
        self.range.ClearContents()

    def column_width(self, w):
        """
        change width of the column(s) of range
        :param w: width
        :return:
        """
        self.range.EntireColumn.ColumnWidth=w

    def row_height(self, h):
        """
        change height of the row(s) of range
        :param h: height
        :return:
        """
        self.range.EntireRow.RowHeight=h

    def curr_region(self):
        """
        get range of the current region
        :return: new range object
        """
        temp=self.range.CurrentRegion.GetAddress(0,0)
        return Rng(address=temp, sheet=self.sheet)

    def replace(self, val, repl_with, whole=False):
        """
        within the range, replace val with repl_with
        :param val: value to be looked for
        :param repl_with: value to replace with
        :return:
        """
        self.range.Replace(What=val,Replacemente=repl_with,LookAt=(_pywintypes.xlWhole if whole else _pywintypes.xlPart))
        pass

    def sort(self, key1, order1=None, key2=None, order2=None, key3=None, order3=None, header=True):
        """
        TODO
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
        formats multiple columns (or rows) at once
        :param fmt_dict: dictionary where keys are column headers of the range (i.e. strings in the first row) while
                         values are excel formatting codes
                         fmt_dict keys may also be regular expressions which are then matched against column names
                         (eg .* may be used as wildcard key, i.e. any columns are matched against this)
        :param cw_dict: same fmt_dict but instead of format strings must contain column widths (row heights)
        :param columns: if True iterate over columns, else over rows
        :return:
        """
        import re
        if columns:
            names = self.range.Rows[1].Value[0]
        else:
            names = list(zip(*self.range.Columns[1].Value))[0]

        instr = ''
        for k, v in fmt_dict.items():
            matcher = re.compile(k)
            for nm in names:
                if matcher.search(nm) is not None:
                    idx = names.index(nm) + 1
                    if columns:
                        self.range.Columns[idx].NumberFormat=v
                    else:
                        self.range.Rows[idx].NumberFormat = v
        for k, v in cw_dict.items():
            matcher = re.compile(k)
            for nm in names:
                if matcher.search(nm) is not None:
                    idx = names.index(nm) + 1
                    if columns:
                        self.range.Columns[idx].ColumnWidth=v
                    else:
                        self.range.Rows[idx].RowHeight = v

    def freeze_panes(self):
        """
        freezes panes at upper left cell of range
        :return:
        """
        self.range.Select()
        self.sheet.workbook.app.ActiveWindow.FreezePanes = True

    def color_scale(self, vmin=5, vmed=50, vmax=95, cv=5):
        """
        TODO
        apply 3 color scale formatting

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
        return

        self.select()
        macro = '"ColorScale" arg1 %f arg2 %f arg3 %f arg4 %i' % (vmin, vmed, vmax, cv)
        ascript = '''
        run XLM macro %s
        ''' % macro
        return _asrun(ascript)

    def select(self):
        """
        select the range
        :return:
        """
        self.range.Select()

    def highlight(self, condition='==', threshold=0.0, interiorcolor=(255, 0, 0)):
        """
        TODO
        highlight cells satisfying given condition
        :param condition: one of ==,!=,>,<,>=,<=
        :param threshold: a number
        :param interiorcolor: an RGB triple specifying the color, eg [255,0,0] is red
        :return:
        """
        return

        dest = self._build_dest()
        ascript = '''
        %s
        tell rng
            try
                delete (every format condition)
            end try
            set newFormatCondition to make new format condition at end with properties {format condition type: cell value, condition operator:operator %s, formula1:%f}
            set color of interior object of newFormatCondition to {%s}
        end tell
        '''
        cond = {'==': 'equal', '!=': 'not equal', '>': 'greater', '<': 'less', '>=': 'greater equal',
                '<=': 'less equal'}
        ascript = ascript % (dest, cond[condition], threshold, str(interiorcolor)[1:-1])
        return _asrun(ascript)

    def col_dict(self):
        """
        given a range with a header row,
        return a dictionary of range objects, each representing a column of the current range
        :return: dict, where keys are header strings, while values are column range objects
        """
        out = {}
        hdr = self.row(1).value()[0]
        c1, r1, c2, r2 = _a2cr(self.address)
        for n, c in zip(hdr, range(c1, c2 + 1)):
            na = _cr2a(c, r1 + 1, c, r2)
            out[n] = Rng(address=na, sheet=self.sheet)
        return out

    def autofit_rows(self):
        """
        autofit row height
        :return:
        """
        self.range.EntireRow.AutoFit()

    def autofit_cols(self):
        """
        autofit column width
        :return:
        """
        self.range.EntireColumn.AutoFit()

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
        self.sheet.ws.Activate()
        self.range.Activate()

    def propagate_format(self, col=True):
        """
        TODO
        propagate formatting of first column (row) to subsequent columns (rows)
        :param col: True for columns, False for rows
        :return:
        """
        return

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
        self.range.Font.Bold = bold
        self.range.Font.Italic = italic
        self.range.Font.Name = name
        self.range.Font.Size = size
        self.range.Font.Color = _rgb2xlcol(color)

    def outline(self, boundaries):
        """
        group rows as defined by boundaries object
        :param boundaries: dictionary, where keys are group "main level" and values is a list of two
                           identifying subrows referring to main level
        :return:
        """
        
        self.sheet.ws.Outline.SummaryRow = 0
        for k, [f, l] in boundaries.items():
            r = self.offset(r=f).resize(r=l - f + 1).entire_row()
            r.range.Group()
            r = self.offset(r=k).row(1)
            r.range.Font.Bold=True

    def show_levels(self, n=2):
        """
        set level of outline to show
        :param n:
        :return:
        """
        self.sheet.ws.Outline.ShowLevels(RowLevels=n)

    def goal_seek(self, target, r_mod):
        """
        TODO
        set value of range of self to target by changing range r_mod
        :param target: the target value for current range
        :param r_mod: the range to modify, or an integer with the column offset from the current cell
        :return:
        """
        return

        dest = self._build_dest()
        if isinstance(r_mod, int):
            dest2 = self.offset(0, r_mod)._build_dest('rng2')
        else:
            dest2 = self.arng(r_mod)._build_dest('rng2')
        ascript = """
        %s
        %s
        goal seek rng goal %f changing cell rng2
        """ % (dest, dest2, target)
        return _asrun(ascript)

    def paste_fig(self, figpath, w, h):
        """
        paste a figure from a file onto an excel sheet, setting width and height as specified
        location will be the top left corner of the current range

        :param figpath: posix path of file
        :param w: width in pixels
        :param h: height in pixels
        :return:
        """

        obj1 = self.sheet.ws.Pictures().Insert(figpath)
        #obj1.ShapeRange.LockAspectRatio = _pywintypes.msoTrue
        obj1.Left = self.range.Left
        obj1.Top = self.range.Top
        obj1.ShapeRange.Width = w
        obj1.ShapeRange.Height = h
        #obj1.Placement = 1
        #obj1.PrintObject = True

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
        TODO

        :param groupby:
        :param totals:
        :param aggfunc:
        :return:
        """
        return

        funcs = ['sum', 'count', 'average', 'maximum', 'minimum', 'product', 'standard deviation']
        assert aggfunc in funcs, "aggfunc must be in " + str(funcs)
        dest = self._build_dest()
        ascript = '''
        %s
        set r1 to value of row 1 of rng
        return my flatten(r1)
        ''' % dest
        names = _parse_aslist(_asrun(ascript))

        igroupby = names.index(groupby) + 1
        itotals = [str(names.index(t) + 1) for t in totals]

        ascript = '''
        %s
        subtotal rng group by %i function do %s total list {%s} summary below data summary above
        ''' % (dest, igroupby, aggfunc, ','.join(itotals))
        return _asrun(ascript)

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
        self.app=_win.gencache.EnsureDispatch('Excel.Application')
        self.app.Visible = True
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
        self.workbooks = []
        for wb in self.app.Workbooks:
            wbo=Workbook(existing=wb,parent=self,name=wb.Name)

    def active_workbook(self):
        """
        return the active workbook
        :return:
        """
        try:
            wb=self.app.ActiveWorkbook
            return Workbook(existing=wb,parent=self,name=wb.Name)
        except:
            raise Exception('no workbook currently open')

    def active_range(self):
        """
        return the active range
        :return:
        """
        try:
            wb = self.app.ActiveWorkbook
            ws = self.app.ActiveSheet
            r = self.app.Selection
            wbo=self.get_wb(wb.Name)
            wso=wbo.get_sheet(ws.Name)
            return Rng(address=r.GetAddress(0,0), sheet=wso)
        except:
            raise Exception('no workbook currently open')

    def create_wb(self, name='Workbook.xlsx'):
        """
        create a new workbook
        :return:
        """
        return Workbook(parent=self)

    def get_wb(self, name):
        """
        get a reference to a workbook based on its name
        :param name:
        :return:
        """
        for i, wb in enumerate(self.workbooks):
            if wb.name == name:
                break
        if len(self.workbooks) == 0 or i > len(self.workbooks):
            raise Exception("there is no workbook %s" % name)
        else:
            return wb

    def calculation(self, manual=True):
        """
        set calculation and screenupdating of excel to manual or automatic
        :param manual:
        :return:
        """
        self.app.Calculation = _win.constants.xlManual if manual else _win.constants.xlAutomatic

    def open_wb(self, fpath):
        """
        open a workbook given its path
        :param fpath:
        :return:
        """
        # self.app.Workbooks.Open(fpath)
        # wb = Workbook(existing=_os.path.basename(fpath), parent=self)
        wb = Workbook(existing=self.app.Workbooks.Open(fpath), parent=self)

        wb.refresh_sheet_list()
        return wb


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

    def __init__(self, existing=None, parent=None, name=None):

        self.wb=None
        self.name = None
        self.parent = None
        self.sheets = []
        self.existing_wb = existing

        if parent is not None:
            self.parent = parent
        elif self.parent is None:
            self.parent = Excel()

        if existing is None:
            wb=self.parent.app.Workbooks.Add()
            if name is not None: wb.Name=name
            self.wb=wb
            self.name = wb.Name
        else:
            self.wb=existing
            self.name = existing.Name
        self.refresh_sheet_list()
        self.parent.workbooks.append(self)

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
        self.sheets=[]
        for ws in self.parent.app.Worksheets:
            Sheet(existing=ws, workbook=self)

    def saveas(self, fpath):
        """
        save a workbook into a different file (silently overwrites existing file with same name!!!)
        :param fpath:
        :return:
        """
        self.parent.app.DisplayAlerts = False
        self.wb.SaveAs(fpath)
        self.parent.app.DisplayAlerts = True
        self.name = _os.path.basename(fpath)

    def save(self, fpath=None):
        """
        save workbook
        :param fpath:
        :return:
        """
        if self.existing_wb is None:
            self.saveas(fpath=fpath)
        else:
            self.parent.app.DisplayAlerts = False
            self.wb.Save()
            self.parent.app.DisplayAlerts = True

    def close(self):
        """
        close a workbook without saving it
        :return:
        """
        self.wb.Close()
        self.parent.refresh_workbook_list()

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

    def __init__(self, existing=None, workbook=None, name=None):


        self.workbook = None
        self.ws = None #reference to the actual worksheet object
        self.name = None
        self.rng = None
        self.cell_data = {}
        self.cell_formats = {}

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
            ws = self.workbook.wb.Worksheets.Add()
            ws.Name = uname

            self.ws=ws
            self.name = self.ws.Name
        else:
            self.ws = existing
            self.name =self.ws.Name
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
        uname = name
        slist = [ws.name for ws in self.workbook.sheets]
        i = 0
        while uname in slist:
            i += 1
            uname = name + '(%i)' % i
        self.ws.Name = uname
        self.name = uname

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
    #     pass

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

    def copy_ws_to_wb(self,target_wb,after=True,sheet_num=0):
        from copy import copy
        if after:
            self.ws.Copy(After=target_wb.get_sheet(sheet_num).ws)
        else:
            self.ws.Copy(Before=target_wb.get_sheet(sheet_num).ws)
        
            
            
def _rgb2xlcol(rgb):
    """
    converts an rgb tuple into a color index as expected by excel
    :param rgb:
    :return:
    """
    strValue = '%02x%02x%02x' % tuple(rgb)
    iValue = int(strValue, 16)
    return iValue

def _parse_windates(v):
    """
    takes a value, a list or a list of list and replaces any pywintypes dates into datetime objects
    :param v:
    :return:
    """
    if isinstance(v,(list,tuple)):
        out=list(v)
        for i in range(len(v)):
            out[i] =_parse_windates(v[i])
    else:
        out=v
        if isinstance(out, _pywintypes.TimeType):
            out = _datetime.datetime(v.year, v.month, v.day, v.hour, v.minute, v.second)
    return out

def _dt2pywintime(d):
    tz=_pytz.timezone('utc')
    if isinstance(d,_datetime.datetime):
        out = _pywintypes.TimeType(d.year, d.month, d.day, d.hour, d.minute, d.second, tzinfo=tz)
    elif isinstance(d,_datetime.date):
        out=_pywintypes.TimeType(d.year,d.month,d.day,tzinfo=tz)
    else:
        out=d
    return out

def _fix_4_win(ll):
    if isinstance(ll,(list,tuple)):
        out=list(ll)
        for i in range(len(ll)):
            out[i]=_fix_4_win(out[i])
    else:
        if isinstance(ll,(_datetime.datetime,_datetime.date)):
            out=_dt2pywintime(ll)
        else:
            out=ll
    return out