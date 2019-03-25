#!/usr/bin/env python
# -*- coding: utf-8 -*-
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

import subprocess as _subprocess #adding the subscore hides the import so that the module namespace remains clean
import pandas as _pd
import os as _os
import datetime as _datetime
import numpy as _np
from copy import copy as _copy
from pyXL import excelpath as _excelpath
from pyXL.excel_utils import _cr2a, _a2cr, _n2x, _x2n, _splitaddr, _df2outline, _isnumeric, _df_to_ll

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
                c=_n2x(col)
                self.address = '%s:%s'%(c,c)
            elif row is not None and col is None:
                self.address = '%i:%i'%(row,row)
            else:
                self.address = '%s%i' % (_n2x(col), row)
            if sheet is not None:
                self.sheet = sheet
        self._set_address()

    def _set_address(self):
        if self.sheet is None:
            self.sheet=Sheet()
        if self.address is None:
            ascript = """
            set addr to get address of selection
            set sn to get name of active sheet
            set wbn to get name of active workbook
            return my list2string({addr,sn,wbn})
            """
            addr = _asrun(ascript)
            temp = addr[1:-1].split('|')
            if self.address is None:
                self.address = temp[0].replace('$', '').replace('"', '')
            if self.sheet is None:
                self.sheet = Sheet(existing=temp[1])

    def _build_dest(self, name='rng'):
        dest='''set %s to range "%s" of worksheet "%s" of workbook "%s"'''
        return dest%(name, self.address, self.sheet.name, self.sheet.workbook.name)

    def arng(self,address=None,row=None,col=None):
        """
        method for quick access to different range on same sheet 
        :param address: 
        :param row: 1-based
        :param col: 1-based
        :return: 
        """
        return Rng(address=address, sheet=self.sheet, row=row,col=col)

    def offset(self, r=0, c=0):
        """
        return new range object offset from the original by r rows and c columns
        :param r: number of rows to offset by
        :param c: number of columns to offset by
        :return: new range object
        """
        coords = _a2cr(self.address)
        if len(coords) == 2:
            newaddr = _cr2a(coords[0]+c, coords[1]+r)
        else:
            newaddr = _cr2a(coords[0] + c, coords[1] + r, coords[2]+c, coords[3]+r)
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
        coords=_a2cr(self.address)
        if len(coords)==2: coords=coords+coords
        if abs:
            newaddr=_cr2a(coords[0], coords[1],coords[0]+max(0,c-1),coords[1]+max(0,r-1))
        else:
            newaddr = _cr2a(coords[0], coords[1], max(coords[0],coords[2] + c), max(coords[1],coords[3] + r))
        return Rng(address=newaddr,sheet=self.sheet)

    def row(self, idx):
        """
        range with given row of current range
        :param idx: indexing is 1-based, negative indices start from last row
        :return: new range object
        """
        coords=_a2cr(self.address)
        if len(coords)==2:
            return _copy(self)
        else:
            newcoords=_copy(coords)
            if idx<0:
                newcoords[1]=newcoords[3]+idx+1
            else:
                newcoords[1]+=idx-1
            newcoords[3]=newcoords[1]
            newaddr=_cr2a(*newcoords)
            return Rng(address=newaddr,sheet=self.sheet)

    def column(self, idx):
        """
        range with given col of current range
        :param idx: indexing is 1-based, negative indices start from last col
        :return: new range object
        """
        coords=_a2cr(self.address)
        if len(coords)==2:
            return _copy(self)
        else:
            newcoords=_copy(coords)
            if idx<0:
                newcoords[0]=newcoords[2]+idx+1
            else:
                newcoords[0]+=idx-1
            newcoords[2]=newcoords[0]
            newaddr=_cr2a(*newcoords)
            return Rng(address=newaddr,sheet=self.sheet)

    def format(self, fmt=None, halignment=None, valignment=None, wrap_text=None):
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
        dest=self._build_dest()
        if fmt is None and halignment is None and valignment is None and wrap_text is None:
            ascript='''
            %s
            return number format of rng
            '''%dest
        else:
            additional=''
            if fmt is not None:
                additional += 'set number format of rng to "%s"\n'%fmt.replace('"','\\"')
            if halignment is not None:
                additional += 'set horizontal alignment of rng to horizontal align %s\n'%halignment
            if valignment is not None:
                additional += 'set vertical alignment of rng to vertical alignment %s\n'%valignment
            if wrap_text is not None:
                additional +='set wrap text of rng to %s\n'%('true' if wrap_text else 'false')
            ascript='''
            %s
            %s
            '''%(dest,additional)
        return _asrun(ascript)

    def filldown(self):
        """
        fill down content of first row to rest of selection
        :return:
        """
        dest=self._build_dest()
        ascript='''
        %s
        fill down rng
        '''%dest
        return _asrun(ascript)

    def color(self, col=None):
        """
        colors a range interior
        :param col: RGB triplet, or None to remove coloring
        :return:
        """
        dest=self._build_dest()
        if col is None:
            ins="set color index of interior object of rng to color index none"
        else:
            ins="set color of interior object of rng to {%s}"%str(col)[1:-1]
        ascript='''
        %s
        %s
        '''%(dest,ins)
        return _asrun(ascript)

    def value(self, v=None):
        """
        get or set the value of a range
        :param v: value to be set, if None current value is returned
        :return:
        """
        dest = self._build_dest()
        if v is not None:
            if _isnumeric(v):
                v = str(v)
            elif isinstance(v,(_datetime.date,_datetime.datetime)):
                v=v.strftime(' date "%A %d %B %Y at %H:%M:%S"')
            elif isinstance(v, str):
                v = '"%s"' % v
            elif isinstance(v, (_pd.DataFrame, _pd.Series)):
                return self.from_pandas(v)
            elif isinstance(v,(list,tuple,_np.ndarray)):
                temp=_pd.DataFrame(v)
                return self.from_pandas(temp,header=False,index=False)
            else:
                raise Exception('Unhandled datatype')
            ascript = '''
            %s
            set value of rng to %s
            ''' % (dest, str(v))
            return _asrun(ascript)
        else:
            ascript = '''
            %s
            set res to value of rng
            return {class of res,res}
            ''' % (dest)
            temp=_asrun(ascript)
            temp=temp[1:-1]
            dtype=temp[:temp.index(', ')]
            val=temp[temp.index(', ')+2:]
            if  dtype=='list':
                return _parse_aslist(val)
            elif dtype=='date':
                return _datetime.datetime.strptime(val, 'date "%A, %d %B %Y at %H:%M:%S"')
            else:
                return eval(val)

    def get_array(self, string_value=False):
        """
        get an excel range as a list of lists

        currently has problems with strings containing the { or } characters
        :return: list
        """
        dest = self._build_dest()
        ascript = '''
        %s
        return %s value of rng
        ''' % (dest,'string' if string_value else '')
        temp = _asrun(ascript)
        temp=_parse_aslist(temp,parse_dates=not string_value)
        return temp

    def to_pandas(self, index=1, header=1):
        """
        return a range as dataframe
        :param index: None for no index, otherwise an integer specifying first n columns to use as index
        :param header: None to avoid using columns, any other value to use first n rows as column header
        :return:
        """
        temp=self.get_array()
        if header is None:
            temp = _pd.DataFrame(temp)
        elif header==1:
            hdr = temp[0]
            temp = _pd.DataFrame(temp[header:], columns=hdr)
        elif header>1:
            hdr=_pd.MultiIndex.from_tuples(temp[:header])
            temp = _pd.DataFrame(temp[header:], columns=hdr)
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
    #     dest = self._build_dest()
    #     ascript='''
    #     %s
    #     '''%dest
    #     if value is not None:
    #         if _U.isnumeric(value):
    #             value = str(value)
    #         elif isinstance(value, _six.string_types):
    #             value = '"%s"' % value
    #         elif isinstance(value, (list, tuple, _np.ndarray)):
    #             temp = _pd.DataFrame(value)
    #             return self.from_pandas(temp, header=False, index=False)
    #         else:
    #             raise Exception('Unhandled datatype')
    #         ascript += '''
    #         set value of rng to %s
    #         ''' % str(value)
    #         return _asrun(ascript)
    #     if formula is not None:
    #         ascript += '''
    #         set formula %s of rng to "%s"
    #         ''' % ('array' if asarray else '',formula.replace('"','\\"'))
    #     if format is not None:
    #         ascript += '''
    #         set number format of rng to "%s"
    #         ''' % format.replace('"','\\"')
    #
    #     if value is None and formula is None and format is None:
    #         ascript += '''
    #         set rng to first cell of rng
    #         return my list2string({class of value of rng, value of rng, string value of rng, has formula of rng, formula of rng, number format of rng})
    #         '''
    #         temp=_asrun(ascript)
    #         temp=[x.strip() for x in temp[1:-1].split('|')]
    #         dtype=temp[0]
    #         val=temp[1]
    #         if dtype=='real':
    #             val = float(val)
    #         else:
    #             val = val
    #         if temp[3]!='true':
    #             formula=''
    #         else:
    #             formula=temp[4]
    #         return val, temp[2], formula, temp[5]
    #     else:
    #         temp = _asrun(ascript)
    #         return temp

    def formula(self, f=None, asarray=False):
        """
        get or set the value of a range
        :param f: formula to be set, if None current formula is returned
        :param asarray: set array formula
        :return:
        """
        dest = self._build_dest()
        if f is not None:
            ascript = '''
            %s
            set formula %s of rng to "%s"
            ''' % (dest, 'array' if asarray else '', f.replace('"','\\"'))
            return _asrun(ascript)
        else:
            ascript = '''
            %s
            set res to formula of rng
            return res
            ''' % (dest)
            temp=_asrun(ascript)
            return temp

    def get_selection(self):
        """
        refresh range coordinates based on current selection
        this modifies the current instance of the object!
        :return:
        """
        self.address=None
        self._set_address()

    def get_cells(self):
        """
        return a list of all addresses of cells in range
        :return:
        """
        if self.size()==(1,1):
            return [self.address]
        else:
            dest = self._build_dest()
            script = '''
            %s
            get address of cells of rng
            ''' % dest
            temp=_asrun(script)
            temp=_parse_aslist(temp)
            return temp

    def from_pandas(self, pdobj, header=True, index=True, index_label=None, outline_string=None):
        """
                write a pandas object to excel
        :param pdobj: any DataFrame or Series object

        #see DataFrame.to_clipboard? for info on params below

        :param header: if False, strip header
        :param index: if False, strip index
        :param index_label: index header
        :param outline_string: a string used to identify outline main levels (eg " All")
        :return:
        """
        crt_size = 5000
        i=0
        if pdobj.ndim==1: pdobj=pdobj.to_frame()
        nrows=int(crt_size/pdobj.shape[1])
        while i<pdobj.shape[0]:
            subpdobj=pdobj.iloc[i:i+nrows,:]
            temp = _df_to_ll(subpdobj,header=header if i==0 else False, index=index)
            asll = _pylist2as(temp)
            trange = self.resize(len(temp), len(temp[0]))
            if i>0: trange=trange.offset(r=i+1,c=0)
            dest = trange._build_dest()
            script = '''
            %s
            set value of rng to %s
            ''' % (dest,asll)
            _asrun(script)
            i+=nrows

        script = '''
        %s
        get address of current region of range "%s"
        ''' % (dest,self.address)
        temp=_asrun(script)
        temp=temp.replace('$','').replace('"','')
        self.address=temp
        if outline_string is not None:
            boundaries=_df2outline(pdobj,outline_string)
            self.outline(boundaries)
        return temp

    def clear_formats(self):
        """
        clear all formatting from range
        :return:
        """
        dest = self._build_dest()
        ascript = '''
        %s
        clear range formats rng
        ''' % dest
        return _asrun(ascript)

    def delete(self, r=None, c=None):
        """
        delete entire rows or columns
        :param r: a (list of) row(s)
        :param c: a (list of) column(s)
        :return:
        """
        assert (r is None) ^ (c is None), "Either r or c must be specified, not both!"
        dest = self._build_dest()
        d=''
        if r is not None:
            if not isinstance(r, (tuple, list)): r = [r]
            for rr in r:
                d += 'delete range row %i of rng shift shift up\n' % rr
        if c is not None:
            if not isinstance(c, (tuple, list)):c=[c]
            for rr in c:
                d += 'delete range column %i of rng shift shift to left\n' % rr
        ascript = '''
        %s
        %s
        ''' % (dest, d)
        return _asrun(ascript)

    def insert(self, r=None, c=None):
        """
        insert rows or columns
        :param r: a (list of) row(s)
        :param c: a (list of) column(s)
        :return:
        """
        assert (r is None) ^ (c is None), "Either r or c must be specified, not both!"
        dest = self._build_dest()
        d=''
        if r is not None:
            if not isinstance(r, (tuple, list)): r = [r]
            for rr in r:
                d += 'insert into range entire row of row %i of rng\n' % rr
        if c is not None:
            if not isinstance(c, (tuple, list)): c = [c]
            for rr in c:
                d += 'insert into range entire column of column %i of rng\n' % rr
        ascript = '''
        %s
        %s
        ''' % (dest, d)
        return _asrun(ascript)

    def clear_values(self):
        """
        clear all values from range
        :return:
        """
        dest = self._build_dest()
        ascript = '''
        %s
        clear contents rng
        ''' % dest
        return _asrun(ascript)

    def column_width(self, w):
        """
        change width of the column(s) of range
        :param w: width
        :return:
        """
        dest = self._build_dest()
        ascript = '''
        %s
        set column width of rng to %s
        ''' % (dest,w)
        return _asrun(ascript)

    def row_height(self, h):
        """
        change height of the row(s) of range
        :param h: height
        :return:
        """
        dest = self._build_dest()
        ascript = '''
        %s
        set row height of rng to %s
        ''' % (dest,h)
        return _asrun(ascript)

    def curr_region(self):
        """
        get range of the current region
        :return: new range object
        """
        dest = self._build_dest()
        ascript = '''
        %s
        get address of current region of rng
        ''' % (dest)
        temp=_asrun(ascript)
        return Rng(address=temp,sheet=self.sheet)

    def replace(self, val, repl_with, whole=False):
        """
        within the range, replace val with repl_with
        :param val: value to be looked for
        :param repl_with: value to replace with
        :return:
        """
        dest = self._build_dest()
        ascript = '''
        %s
        replace rng what "%s" replacement "%s" %s
        ''' % (dest,val,repl_with, 'look at whole' if whole else 'look at part')
        return _asrun(ascript)

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
        dest = self._build_dest()
        ascript = '''
        %s
        set r1 to row 1 of rng
        set f to cell 1 of r1
        set header to value of r1
        set header to item 1 of header
        '''%dest
        kd=''
        for i,k,d in list(zip(range(3),[key1,key2,key3],[order1,order2,order3])):
            if k is not None:
                if d is None: d='a'
                ascript+="""
                set idx_%i to my get_index_in_list("%s", header)
                set offs_%i to get offset f column offset (idx_%i - 1)
                set offs_%i to get address of offs_%i
                """%(i,k,i,i,i,i)
                kd+=""" key%i offs_%i order%i sort %s """%(i+1,i,i+1,'ascending' if d.lower()[0]=='a' else 'descending')
        ascript+='''
        activate object worksheet object of rng
        sort rng %s %s
        '''%(kd,'with header' if header else '')

        return _asrun(ascript)

    def size(self):
        """
        return size of range
        :return: columns, rows
        """
        temp=_a2cr(self.address)
        if len(temp)==2:
            return (1,1)
        return temp[2]-temp[0]+1,temp[3]-temp[1]+1

    def coords(self):
        """
        return coordinates of range
        :return: left,top,right,bottom
        """
        temp=_a2cr(self.address)
        if len(temp)==2:
            return temp[0],temp[1],temp[0],temp[1]
        return temp[0],temp[1],temp[2],temp[3]

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
        dest = self._build_dest()
        ascript = '''
        %s
        set r1 to value of %s 1 of rng
        return my flatten(r1)
        ''' % (dest,'row' if columns else 'column')
        names =_parse_aslist(_asrun(ascript))

        instr=''
        for k, v in fmt_dict.items():
            matcher = re.compile(k)
            for nm in names:
                if matcher.search(nm) is not None:
                    idx=names.index(nm) + 1
                    instr+='set number format of %s %i of rng to "%s"\n'%('column' if columns else 'row',idx,v.replace('"','\\"'))
        for k, v in cw_dict.items():
            matcher = re.compile(k)
            for nm in names:
                if matcher.search(nm) is not None:
                    idx=names.index(nm) + 1
                    instr += 'set %s of column %i of rng to %f\n' % ('column width' if columns else 'row height',idx, v)
        ascript = '''
        %s
        %s
        ''' % (dest,instr)
        return _asrun(ascript)

    def freeze_panes(self):
        """
        freezes panes at upper left cell of range
        :return:
        """
        dest = self._build_dest()
        ascript = '''
        %s
        select rng
        set freeze panes of active window to true
        ''' % dest
        return _asrun(ascript)

    def color_scale(self, vmin=5, vmed=50, vmax=95, cv=5, invert_colors=False):
        """
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
        '''

        '''
        self.select()
        macro = '"ColorScale" arg1 %f arg2 %f arg3 %f arg4 %i arg5 %s' % (vmin, vmed, vmax, cv, str(invert_colors))
        ascript = '''
        run XLM macro %s
        ''' % macro
        return _asrun(ascript)

    def select(self):
        """
        select the range
        :return:
        """
        dest = self._build_dest()
        ascript = '''
        %s
        select rng
        ''' % dest
        return _asrun(ascript)

    def highlight(self, condition='==', threshold=0.0, interiorcolor=(255, 0, 0)):
        """
        highlight cells satisfying given condition
        :param condition: one of ==,!=,>,<,>=,<=
        :param threshold: a number
        :param interiorcolor: an RGB triple specifying the color, eg [255,0,0] is red
        :return:
        """
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
        cond={'==': 'equal', '!=': 'not equal', '>': 'greater', '<': 'less', '>=': 'greater equal', '<=': 'less equal'}
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
        for n, c in zip(hdr, range(c1, c2+1)):
            na = _cr2a(c, r1+1, c, r2)
            out[n] = Rng(address=na, sheet=self.sheet)
        return out

    def autofit_rows(self):
        """
        autofit row height
        :return:
        """
        dest = self._build_dest()
        ascript = '''
        %s
        autofit every row of rng
        ''' % dest
        return _asrun(ascript)

    def autofit_cols(self):
        """
        autofit column width
        :return:
        """
        dest = self._build_dest()
        ascript = '''
        %s
        autofit every column of rng
        ''' % dest
        return _asrun(ascript)

    def entire_row(self):
        """
        get entire row(s) of current range
        :return: new object
        """
        c=_a2cr(self.address)
        if len(c)==2: c+=c
        cc='%s:%s'%(c[1],c[3])
        return Rng(address=cc,sheet=self.sheet)

    def entire_col(self):
        """
        get entire row(s) of current range
        :return: new object
        """
        c=_a2cr(self.address)
        if len(c)==2: c+=c
        cc='%s:%s'%(_n2x(c[0]),_n2x(c[2]))
        return Rng(address=cc,sheet=self.sheet)

    def activate(self):
        """
        activate range
        :return:
        """
        dest = self._build_dest()
        ascript = '''
        %s
        activate object worksheet "%s"
        select rng
        ''' % (dest,self.sheet.name)
        return _asrun(ascript)

    def propagate_format(self, col=True):
        """
        propagate formatting of first column (row) to subsequent columns (rows)
        :param col: True for columns, False for rows
        :return:
        """
        dest = self._build_dest()
        ascript='''
        %s
        set r to (get %s in rng)
        copy range item 1 of r
        repeat with i from 2 to length of r
            paste special item i of r what paste formats
        end repeat
        '''%(dest,'columns' if col else 'rows')
        return _asrun(ascript)

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
        dest = self._build_dest()
        ascript='''
        %s
        set properties of font object of rng to {bold:%s, italic:%s, name:"%s", font size:%s, color:{%s}}
        return properties of font object of rng
        ''' % (dest, bold, italic, name, size, str(color)[1:-1])
        return _asrun(ascript)

    def outline(self, boundaries):
        """
        group rows as defined by boundaries object
        :param boundaries: dictionary, where keys are group "main level" and values is a list of two
                           identifying subrows referring to main level
        :return:
        """
        dest = self._build_dest()
        scr="""
        %s
        set summary row of outline object of worksheet object of rng to summary above
        """%dest
        for k,[f, l] in boundaries.items():
            r=self.offset(r=f).resize(r=l-f+1).entire_row()
            dest = r._build_dest()
            scr+="""
            %s
            group entire row of rng
            """%dest
            r=self.offset(r=k).row(1)
            dest = r._build_dest()
            scr+="""
            %s
            set bold of font object of rng to true
            """%dest

        return _asrun(scr)

    def show_levels(self, n=2):
        """
        set level of outline to show
        :param n:
        :return:
        """
        dest = self._build_dest()
        ascript="""
        %s
        show levels outline object of worksheet object of rng row levels %i
        """%(dest,n)
        return _asrun(ascript)

    def goal_seek(self, target, r_mod):
        """
        set value of range of self to target by changing range r_mod
        :param target: the target value for current range
        :param r_mod: the range to modify, or an integer with the column offset from the current cell
        :return:
        """
        dest = self._build_dest()
        if isinstance(r_mod,int):
            dest2 = self.offset(0,r_mod)._build_dest('rng2')
        else:
            dest2 = r_mod._build_dest('rng2')
        ascript="""
        %s
        %s
        goal seek rng goal %f changing cell rng2
        """%(dest,dest2,target)
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
        dest = self._build_dest()
        ascript="""
        %s
        set ws to worksheet object of rng
        set theFilePath to Posix file "%s"
        set newPic to make new picture at the beginning of ws with properties {file name:theFilePath}
        set left position of newPic to left position of rng
        set top of newPic to top of rng
        set width of newPic to %i
        set height of newPic to %i
        """%(dest,figpath,w,h)
        return _asrun(ascript)

    def subrng(self, t, l, nr=1, nc=1):
        """
        given a range returns a subrange defined by relative coordinates
        :param t: row offset from current top row
        :param l: column offset from current top column
        :param nr: number of rows in subrange
        :param nc: number of columns in subrange
        :return: range object
        """
        coords=_a2cr(self.address)
        newaddr = _cr2a(coords[0]+l, coords[1]+t, coords[0]+l+nc-1, coords[1]+t+nr-1)
        return Rng(address=newaddr, sheet=self.sheet)

    def subtotal(self,groupby,totals,aggfunc='sum'):
        """
        
        :param groupby: 
        :param totals: 
        :param aggfunc: 
        :return: 
        """
        funcs=['sum','count','average','maximum','minimum','product','standard deviation']
        assert aggfunc in funcs, "aggfunc must be in "+str(funcs)
        dest = self._build_dest()
        ascript = '''
        %s
        set r1 to value of row 1 of rng
        return my flatten(r1)
        ''' % dest
        names = _parse_aslist(_asrun(ascript))

        igroupby=names.index(groupby)+1
        itotals=[str(names.index(t)+1) for t in totals]

        ascript = '''
        %s
        subtotal rng group by %i function do %s total list {%s} summary below data summary above
        ''' % (dest, igroupby, aggfunc, ','.join(itotals))
        return _asrun(ascript)



def _asrun(ascript, application=_excelpath):
    """
    Run the given AppleScript and return the standard output
    raises exception if std error is not empty
    :param ascript: any valid applescript code
    :return:
    """
    ascript="""
    tell application "%s"
    %s
    end tell

    on get_index_in_list(this_item, this_list)
        repeat with i from 1 to the count of this_list
            if item i of this_list is this_item then return i
        end repeat
        return 0
    end get_index_in_list

    on list2string(theList)
        set AppleScript's text item delimiters to "|"
        set theString to theList as string
        set AppleScript's text item delimiters to ""
        return theString
    end list2string

    on flatten(aList)
        if class of aList is not list then
            return {aList}
        else if length of aList is 0 then
            return aList
        else
            return flatten(first item of aList) & (flatten(rest of aList))
        end if
    end flatten
    """ % (application, ascript)
    ascript=bytes(ascript, 'utf-8')
    osa = _subprocess.Popen(['osascript', '-s', 's', '-'],
                            stdin=_subprocess.PIPE,
                            stdout=_subprocess.PIPE,
                            stderr=_subprocess.PIPE)
    stdout, stderr = osa.communicate(ascript)
    res = stdout.decode('utf-8', 'ignore')[:-1]
    err = stderr.decode('utf-8', 'ignore')[:-1]
    if len(err) > 0:
        try:
            print(ascript.decode('utf-8'))
        except:
            print(ascript.decode('ascii',errors='ignore'))
        raise Exception(err)
    else:
        return res

class Excel():
    """
    basic wrapper of Excel application, providing some methods to perform simple automation, such as
    creating/opening/closing workbooks
    Also, keeps count of open workbooks and specific application settings, such as calculation manual etc.
    """
    def __repr__(self):
        """
        print out coordinates of the range object
        :return:
        """
        return 'Excel application, currently %i workbooks are open'%len(self.workbooks)

    def __init__(self):
        self.workbooks = []
        self._calculation_manual = False
        self.refresh_workbook_list()
        chk=_check_date_format()
        if not chk:
            print("Warning: system date format is incorrect, thus handling dates might not work correctly")
            print("Please go to System Preferences->Language & Region->Advanced->Dates, and make sure that")
            print("'Full' is set to something like 'Thursday, 5 January 2017' including the comma and spaces!")

    def refresh_workbook_list(self):
        """
        make sure that object is consistent with current state of excel
        this needs to be called if, during an interactive session, users create/delete workbooks manually,
        as the Excel object has no way to know what the user does
        :return: 
        """
        scr="""
        get my list2string(name of workbooks)
        """
        temp=_asrun(scr)
        self.workbooks = []
        if temp!='"missing value"':
            temp=temp[1:-1].split('|')
            for t in temp:
                wb=Workbook(existing=t,parent=self)
                wb.refresh_sheet_list()

    def active_workbook(self):
        """
        return the active workbook
        :return: 
        """
        scr="""
        get name of active workbook
        """
        temp=_asrun(scr)[1:-1]
        if temp=='missing value':
            raise Exception('no workbook currently open')
        else:
            return self.get_wb(temp)

    def active_range(self):
        """
        return the active range
        :return: 
        """
        scr="""
        {name of active workbook, name of active sheet, get address of selection}
        """
        temp=_asrun(scr)
        if temp=='missing value':
            raise Exception('no workbook currently open')
        else:
            temp=_parse_aslist(temp)
            wb=self.get_wb(temp[0])
            return wb.get_sheet(temp[1]).arng(temp[2])

    def create_wb(self,name='Workbook.xlsx'):
        """
        create a new workbook
        :return: 
        """
        return Workbook(parent=self)

    def get_wb(self,name):
        """
        get a reference to a workbook based on its name
        :param name: 
        :return: 
        """
        for i,wb in enumerate(self.workbooks):
            if wb.name==name:
                break
        if len(self.workbooks)==0 or i>len(self.workbooks):
            raise Exception("there is no workbook %s"%name)
        else:
            return wb

    def calculation(self,manual=True):
        """
        set calculation and screenupdating of excel to manual or automatic
        :param manual: 
        :return: 
        """
        ascript='''
        set calculation to calculation %s
        set screenupdating to %s
        ''' %('manual' if manual else 'automatic',str(manual))
        temp= _asrun(ascript)
        self._calculation_manual=manual

    def open_wb(self,fpath):
        """
        open a workbook given its path
        :param fpath: 
        :return: 
        """
        ascript='''
        open workbook workbook file name POSIX file "%s"
        ''' %_os.path.abspath(_os.path.expanduser(fpath))
        temp= _asrun(ascript)
        wb = Workbook(existing=_os.path.basename(fpath),parent=self)
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
        return "Workbook object '%s', has %i sheets" % (self.name,len(self.sheets))

    def __init__(self,existing=None,parent=None, name=None):

        self.name = None
        self.parent = None
        self.sheets = []

        if parent is not None:
            self.parent=parent
        elif self.parent is None:
            self.parent=Excel()

        if existing is None:
            ascript = '''
                set wb to make new workbook
                return {name of wb, name of first sheet of wb}
                '''
            wb, ws = _parse_aslist(_asrun(ascript))
            self.name = wb
            Sheet(existing=ws, workbook=self)
        else:
            self.name=existing
        self.parent.workbooks.append(self)

    def create_sheet(self,name='Sheet'):
        """
        create a new sheet in the current workbook
        :param name: 
        :return: 
        """
        return Sheet(name=name,workbook=self)

    def refresh_sheet_list(self):
        """
        make sure that object is consistent with current state of excel
        this needs to be called if, during an interactive session, users create/delete sheets manually,
        as the Excel object has no way to know what the user does
        :return: 
        """
        scr="""
        get my list2string(name of sheets of workbook "%s")
        """%self.name
        temp=_asrun(scr)
        if temp!='"missing value"':
            temp=temp[1:-1].split('|')
            self.sheets=[]
            for t in temp:
                Sheet(existing=t,workbook=self)

    def saveas(self,fpath):
        """
        save a workbook into a different file (silently overwrites existing file with same name!!!)
        :param fpath: 
        :return: 
        """
        scr="""
        set fname to (POSIX file "%s") as string
        save workbook as workbook "%s" filename fname overwrite True
        """%(fpath,self.name)
        temp=_asrun(scr)
        self.name=_os.path.basename(fpath)

    def save(self,fpath):
        """
        save workbook
        :param fpath:
        :return:
        """
        scr="""
        save workbook "%s" in "%s"
        """%(self.name,fpath)
        temp=_asrun(scr)
        self.name=_os.path.basename(fpath)

    def close(self):
        """
        close a workbook without saving it
        :return: 
        """
        scr="""
        close workbook "%s" saving no
        """%self.name
        temp=_asrun(scr)
        self.parent.refresh_workbook_list()

    def get_sheet(self,name):
        """
        get a reference to a sheet object given a name
        :param name: 
        :return: 
        """
        # for i,sh in enumerate(self.sheets):
        #     if sh.name==name:
        #         break
        # if len(self.sheets)==0 or i>len(self.sheets):
        #     raise Exception("there is no sheet %s"%name)
        # else:
        #     return sh
        for i,sh in enumerate(self.sheets):
            if sh.name==name:
                break
        else:
            raise Exception("There is no sheet '%s'" % name)
        return  sh



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
        self.name = None
        self.rng = None

        if workbook is None:
            self.workbook=Workbook()
            existing=self.workbook.sheets[0].name
        else:
            self.workbook=workbook


        if existing is None:
            if name is None: name = 'Sheet'
            wb=self.workbook.name
            if name is None: name='Sheet'
            ascript='''
            set wb to workbook "%s"
            set ns to make new sheet at wb
            set i to 0
            set sname to "%s"
            repeat
                try
                    set name of ns to sname
                    exit repeat
                on error
                    set i to i + 1
                    set sname to "%s" & " (" & i & ")"
                end try
            end repeat
            return sname
            '''%(wb, name, name)
            temp=_asrun(ascript)[1:-1]
            self.name=temp
        else:
            self.name=existing
        self.rng=Rng('A1',sheet=self)
        self.workbook.sheets.append(self)

    def arng(self,address=None, row=None, col=None):
        """
        access a range on the sheet, providing either address in A1 format, or a row and/or a column
        :param address: string in A1 format
        :param row: integer, 1-based
        :param col: integer, 1-based
        :return: 
        """
        self.rng=Rng(address=address,row=row,col=col,sheet=self)
        return self.rng

    def rename(self, name):
        """
        change the name of the current sheet
        :param name: 
        :return: 
        """
        ascript='''
        set ns to sheet "%s"
        set i to 0
        set sname to "%s"
        repeat
            try
                set name of ns to sname
                exit repeat
            on error
                set i to i + 1
                set sname to "%s" & " (" & i & ")"
            end try
        end repeat
        return sname
        '''%(self.name,name,name)
        temp=_asrun(ascript)[1:-1]
        self.name=temp

    def delete_shapes(self):
        """
        delete all shape objects on the current sheet
        :return: 
        """
        ascript = '''
        set shlist to get shapes of sheet "%s"
        delete items of shlist
        ''' % self.name
        return _asrun(ascript)

    def unprotect(self):
        """
        remove protection (only if no password!)
        :return:
        """
        ascript = '''
        unprotect sheet "%s"
        ''' % self.name
        return _asrun(ascript)

    def protect(self):
        """
        activate protection (without password!)
        :return:
        """
        ascript = '''
        protect worksheet sheet "%s"
        ''' % self.name
        return _asrun(ascript)

    # def get_values_formulas_formats(self, *rngs):
    #     """
    #     traverses a range and returns its contents as a dictionary
    #     keys are cell addresses, values are content, formulas and formats
    #
    #     :param rngs: one or more range addresses
    #     :return: 3 dicts
    #     """
    #     outvalues = {}
    #     outformulas = {}
    #     outformats = {}
    #     for r in rngs:
    #         rn = Rng(r,sheet=self)
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
        ascript = 'tell sheet "%s" of workbook "%s"\n'%(self.name,self.workbook.name)
        if values_dict is not None:
            for addr, v in values_dict.items():
                if _isnumeric(v):
                    v = str(v)
                elif isinstance(v, str):
                    v = '"%s"' % v
                else:
                    raise Exception('Unhandled datatype')
                ascript += 'set value of range "%s" to %s\n' % (addr, v)
        if formats_dict is not None:
            for addr, v in formats_dict.items():
                ascript += 'set number format of range "%s" to "%s"\n' % (addr, v.replace('"','\\"'))
        if formulas_dict is not None:
            for addr, v in formulas_dict.items():
                ascript += 'set formula of range "%s" to "%s"\n' % (addr, v.replace('"','\\"'))
        if arrformulas_dict is not None:
            for addr, v in arrformulas_dict.items():
                ascript += 'set formula array of range "%s" to "%s"\n' % (addr, v.replace('"','\\"'))
        ascript += 'end tell\n'
        return _asrun(ascript)


def _parse_aslist(aslist, parse_dates=True):
    """
    takes a string representation of an applescript list and tries to transform it into a python list
    date objects are explicitly handled
    :param aslist: 
    :return: 
    """
    if parse_dates:
        try:
            i = aslist.index('date "')
        except:
            i = 0
        while i > 0:
            i2 = aslist[i + 7:].index('"') + (i + 7)
            dd = aslist[i:i2 + 1]
            ddate = _datetime.datetime.strptime(dd, 'date "%A, %d %B %Y at %H:%M:%S"')
            aslist = aslist.replace(dd, '_datetime.'+repr(ddate)[9:])
            try:
                i = aslist.index('date "')
            except:
                i = 0
    # here add some way to make sure that braces within strings are NOT replaced!
    aslist = aslist.replace('{', '[').replace('}', ']')
    aslist=eval(aslist)
    return aslist

def _pylist2as(pylist):

    out='{'
    for el1 in pylist:
        if isinstance(el1,(list,tuple)):
            temp=_pylist2as(el1)
            out+=temp+','
        elif isinstance(el1,str):
            out+='"%s",'%(el1).replace('"','\\"')
        elif isinstance(el1, _datetime.date):
            out+=el1.strftime('date "%A, %d %B %Y at %H:%M:%S"')+','
        else:
            if el1 is None:
                out += "null,"
            elif _np.isnan(el1):
                out += "null,"
            else:
                out+=str(el1)+','
    out=out[:-1]+'}'
    return out

def _check_date_format():
    temp=_asrun("current date")
    try:
        out=_datetime.datetime.strptime(temp, 'date "%A, %d %B %Y at %H:%M:%S"')
        return True
    except:
        return False
