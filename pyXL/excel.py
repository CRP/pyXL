#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys as _sys
import platform as _platform

my_platform=_platform.system()
this = _sys.modules[__name__]
this.engine = None
this.interactive = False
this.Excel = None
this.Rng = None
this.Workbook = None
this.Sheet = None

def switch_engine(engine='xlsxwriter'):
    """
    switch the excel engine, the same routine may use different engines without changes to the code

    For each engine the module exposes the following objects:
    - Excel: an abstract object encapsulating the Excel application session
    - Workbook: an abstract object encapsulating an Excel file
    - Sheet: an abstract object encapsulating a sheet in a Workbook
    - Rng: an abstract object encapsulating a range on a Sheet
    and some functions such as:
    - df2rng: write a pandas object to a range
    - rng2arr: read the content of a range into a list
    - rng2df: read the content of a range into a pandas object

    See the docstrings for these objects for more details.

    :param engine: can be one of the following:
                applescript: new mac excel engine using applescript and discarding the appscript module which does
                    not work in python 3; also, the new engine is object oriented, thus clearer and easier to use
                win32com: new engine to remote control windows excel, same interface as the mac one
                xlsxwriter/file: new engine to create excel files, uses same interface as the mac engine, with a subset
                    of capabilities; specifically, it can not open or read existing files
    :return:
    """

    if engine in ('xlsxwriter','file'):
        import pyXL.excel_xlsxwriter as XLXWR
        this.Excel=XLXWR.Excel
        this.Rng = XLXWR.Rng
        this.Workbook = XLXWR.Workbook
        this.Sheet = XLXWR.Sheet
        this.engine=engine
        this.interactive =False
        print("Switched to XLSXWriter Excel engine")
    elif engine in ('applescript'):
        if my_platform != 'Darwin':
            print("Applescript engine only works on MacOS platforms, falling back to xlsxwriter")
            switch_engine('file')
            return
        import pyXL.excel_mac_as as XLMAC
        this.Excel=XLMAC.Excel
        this.Rng = XLMAC.Rng
        this.Workbook = XLMAC.Workbook
        this.Sheet = XLMAC.Sheet
        this.engine=engine
        this.interactive = True
        print("Switched to Applescript Excel engine")
    elif engine in ('win32com'):
        if my_platform != 'Windows':
            print("win32com engine only works on Windows platforms, falling back to xlsxwriter")
            switch_engine('file')
            return
        import pyXL.excel_win32 as XLWIN
        this.Excel = XLWIN.Excel
        this.Rng = XLWIN.Rng
        this.Workbook = XLWIN.Workbook
        this.Sheet = XLWIN.Sheet
        this.engine = engine
        this.interactive = True
        print("Switched to win32com Excel engine")
    else:
        print("Unknown engine, falling back to xlsxwriter")
        switch_engine('file')

def test():
    """
    routine to test excel functionality
    :return:
    """
    from pandas import DataFrame
    import numpy as np, os
    import matplotlib.pyplot as plt

    print("creating link to Excel application")
    x=this.Excel() # this is just a reference to excel and allows to set a few settings, such as calculation etc.
    print(x)

    print("creating new workbook")
    wb=x.create_wb() # create a new workbook, returns instance of Workbook object
    print(x)

    print(wb)
    print("creating new worksheet")
    sh=wb.create_sheet('pippo') # create a new sheet named "pippo", returns instance of Sheet object
    print(wb)
    print("write someting on sheet")
    sh.arng("A1").value("pluto")

    print("create a random image")
    try:
        rand_array = np.random.rand(550, 550)  # create your array
        ax=plt.imshow(rand_array)
        filename = './testimage.jpg'
        ax.figure.savefig(filename)
        plt.close(ax.figure)
        sh.arng("C3").paste_fig(filename,w=300,h=300)
    except:
        print("Failed random image creation")

    print("now working on initial sheet")
    sh2=wb.sheets[0] # get reference of first sheet
    r=sh2.arng('B2') # access a range on the sheet
    print(r) #prints current "coordinates"

    print("generating a dataframe")
    np.random.seed(1)
    temp=DataFrame(np.random.randn(30,4),columns=list('abcd'))
    temp=temp.groupby([(temp.index / 10).map(lambda x: int(x)), [a % 2 ==0 for a in temp.index],temp.index]).sum()
    print("and writing it onto excel with outline")
    r.from_pandas(temp)#,outline_string=-1) #write data to current sheet
    print("coordinates have changed!!")
    print(r)

    print("do some formatting")
    r.format_range(fmt_dict={'index':'0','b':'@','d':'0.0000','c':'0.0%'}, cw_dict={'c':20})
    r.column(2).font_format(bold=True,color=[255,0,0])

    #print("do some sorting")
    #r.sort('b')

    print("now write a formula")
    r2=r.arng('C40')
    r2.formula("=sqrt(D7*3)")

    print("now write a formula array")
    r3=r.offset(c=8).subrng(t=1,l=1,nr=r.size()[1]-1,nc=r.size()[0]-1)
    r3.formula("=%s*2"%r.subrng(t=1,l=1,nr=r.size()[1]-1,nc=r.size()[0]-1).address,asarray=True)
    print("add a header")
    r.offset(c=8).column(1).value([[None,'aa','bb','cc','dd']])

    breakpoint()
    print("now read back data from formula array range")
    a=r3.curr_region().to_pandas(index=None) #read data from current sheet
    print(a)

    if this.interactive:
        r=sh2.arng("B2:E6")
        r.activate()
        addr=x.active_range().address
        if addr==r.address:
            print("internal address and address of active range returned by excel match!")
        else:
            print("WARNING! internal address and address of active range returned by excel do not match!")

    print("now save and then close excel file")
    wb.saveas("test.xlsx")
    wb.close()

def df2rng(df, rng=None, index_header='', skip_header=False, skip_index=False, sparse_mi=False, outline=None):
    """
    write a dataframe to an excel range
    :param df:
    :param rng: destination range (string like "A1")
    :param index_header: override the name of the index of the df
    :param skip_header:
    :param skip_index:
    :param sparse_mi: not used
    :param outline: create subgroups and outline whenever this string appears in the index
    :return:
    """
    if isinstance(df,dict):
        for addr,sdf in df.items():
            df2rng(sdf,addr,index_header=index_header,skip_header=skip_header,skip_index=skip_index,sparse_mi=sparse_mi,outline=outline)
    else:
        x = this.Excel()
        n=len(x.workbooks)

        if n==0:
            ws = x.create_wb().sheets[0]
            if isinstance(rng,str):
                range=ws.arng(rng)
            else:
                range = ws.arng('A1')
        else:
            r = x.active_range()
            if isinstance(rng,str):
                range=r.sheet.arng(rng)
            else:
                range = r

        range.from_pandas(df, header=not skip_header, index=not skip_index, index_label=index_header, outline_string=outline)

def rng2arr(rng=None, string_value=False, c=False):
    """
    read data from excel active range into list
    :return:
    """
    x = this.Excel()
    r = x.active_range()
    r.get_selection()
    out = r.value()
    if len(out)==1 and c:
        out=out[0]
    elif len(out)>1 and len(out[0])==1 and c:
        out=list(list(zip(*out))[0])
    return out


def arr2rng(arr, rng=None):
    """
    write list onto excel
    :return:
    """
    x = this.Excel()
    if rng is None:
        r = x.active_range()
    else:
        r = x.active_range().arng(rng)
    r.value(arr)

def rng2df(rng=None, first_col_as_index=True,first_row_as_columns=True, dtype=None):
    """
    read excel active range into dataframe
    :return:
    """
    from qlib.tseries import pd as _pd
    x = this.Excel()
    if rng is None:
        r = x.active_range()
        r.get_selection()
    else:
        assert isinstance(rng,str), "rng can only be an Excel range, eg. A1:B2"
        r=x.active_range().arng(rng)
    temp = r.value()
    if first_row_as_columns:
        cols = temp[0]
        vals = temp[1:]
    else:
        cols = None
        vals = temp
    temp = _pd.DataFrame.from_records(list(vals), columns=cols)
    if first_col_as_index:
        temp = temp.set_index(temp.columns[0])
    if dtype is not None:
        temp=temp.astype(dtype)
    return temp

def propagate_format(col=True):
    """
    take the format of the first column (or row) of the current selection, and paste format to each subsequent
        column (row)
    :param col:
    :return:
    """
    x = this.Excel()
    r = x.active_range()
    r.propagate_format(col=col)

def ar():
    """
    quick way to get a reference to current active range in Excel
    :return:
    """
    x = this.Excel()
    r = x.active_range()
    return r

def repeat_goal_seek(target_rng=None, target_values=0, changing_rng=1):
    """
    apply goal seek to each cell in target_rng, changing values in changing_rng in order to converge to target_values
    :param target_rng: a Rng object (should be only one column wide)
    :param target_values: a scalar value, or a list of values, or a range with a list of values
    :param changing_rng: a range or an integer representing column offset from target_rng (-1 means the 1st col on
                        the left side of target_rng)
    :return:
    """
    if target_rng is None: target_rng=ar()
    ws=target_rng.sheet
    if isinstance(changing_rng,int):
        changing_rng=target_rng.offset(c=changing_rng)
    if isinstance(target_values,this.Rng):
        target_values=list(zip(*target_values.value()))[0]
    elif not (isinstance(target_values,list) or isinstance(target_values,tuple)):
        target_values=[target_values]*target_rng.size()[1]
    data=zip(*[target_rng.get_cells(),target_values,changing_rng.get_cells()])
    for tgt_r,tgt_v,chg_r in data:
        print("Try to force %s to %f by changing %s"%(tgt_r,tgt_v,chg_r))
        ws.arng(tgt_r).goal_seek(target=tgt_v,r_mod=ws.arng(chg_r))

def cells2list():
    """
    get values of selected (not necessarily contiguous) cells in a flat list
    :return:
    """
    r=ar()
    addresses=r.address.split(',')
    cells=[]
    for address in addresses:
        tr=r.arng(address)
        cells+=tr.get_cells()
    out=[]
    for cell in cells:
        tr=r.arng(cell)
        out+=[tr.value()]
    return out

def midf2xl(midf, axis=1):
    """
    write a multiindex dataframe to an excel workbook, where a different worksheet is used for each value of
     the multiindices' first level
    :param midf:
    :param axis:
    :return:
    """
    x=this.Excel()
    if axis==0:
        vals=midf.index.levels[0]
        wb=x.create_wb()
        for val in vals:
            ws=wb.create_sheet(val)
            ws.arng("A1").from_pandas(midf.loc[val])
    elif axis==1:
        vals=midf.columns.levels[0]
        wb=x.create_wb()
        for val in vals:
            ws=wb.create_sheet(val)
            ws.arng("A1").from_pandas(midf[val])
    else:
        raise Exception("axis must be 0 or 1")
