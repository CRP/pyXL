#!/usr/bin/env python
# -*- coding: utf-8 -*-

def _cr2a(c1, r1, c2=None, r2=None):
    """
    r1=1 c1=1 gives A1 etc.
    """
    assert r1>0 and c1>0, "negative coordinates not allowed!"
    out=_n2x(c1)+str(r1)
    if c2 is not None:
        out +=':'+ _n2x(c2) + str(r2)
    return out


def _a2cr(a,f4=False):
    """
    B1 gives [1,2]
    B1:D3 gives [1,2,3,4]

    if f4==True, always return a 4-element list, so [1,2] becomes [1,2,1,2]
    """
    if ':' in a:
        tl,br=a.split(':')
        out= _a2cr(tl)+_a2cr(br)
        if out[0]==0:out[0]=1
        if out[1]==0:out[1]=1
        if out[2]==0:out[2]=2**14
        if out[3]==0:out[3]=2**20
        return out
    else:
        c,r=_splitaddr(a)
        if f4:
            return [_x2n(c), r, _x2n(c), r]
        else:
            return [_x2n(c),r]


def _n2x(n):
    """
    convert decimal into base 26 number-character
    :param n:
    :return:
    """
    numerals='ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    b=26
    if n<=b:
        return numerals[n-1]
    else:
        pre=_n2x((n-1)//b)
        return pre+numerals[n%b-1]


def _x2n(x):
    """
    converts base 26 number-char into decimal
    :param x:
    :return:
    """
    numerals = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    b=26
    n=0
    for i,l in enumerate(reversed(x)):
        n+=(numerals.index(l)+1)*b**(i)
    return n


def _splitaddr(addr):
    """
    splits address into character and decimal
    :param addr:
    :return:
    """
    col='';rown=0
    for i in range(len(addr)):
        if addr[i].isdigit():
            col = addr[:i]
            rown = int(addr[i:])
            break
        elif i==len(addr)-1:
            col=addr
    return col,rown

def _df2outline(df, outline_string):
    """
    infer boundaries of a dataframe given an outline_string, to be fed into rng.outline(boundaries)
    :param df:
    :param outline_string:
    :return:
    """
    from collections import OrderedDict
    out=OrderedDict()
    idx = list(zip(*df.index.values.tolist()))
    z=1
    rf = 1  # this is just to avoid an exception if df has no index field named outline
    for lvl in range(len(idx)):
        for i, v, p in list(zip(list(range(1, len(idx[lvl]))), idx[lvl][1:], idx[lvl][:-1])):
            if p == outline_string: rf = i
            if (v == outline_string and p != outline_string) or i + 1 == len(idx[lvl]):
                rl = (i if i + 1 == len(idx[lvl]) else i - 1)
                out[z+rf-1]=[z+rf,z+rl]
    return out

def _isrow(addr):
    if ':' in addr:
        coords=_a2cr(addr)
        return coords[2]-coords[0]==16383
    else: return False

def _iscol(addr):
    if ':' in addr:
        coords=_a2cr(addr)
        return coords[3]-coords[1]==1048575
    else: return False

def _isnumeric(x):
    '''
    returns true if x can be cast to a numeric datatype
    '''
    try:
        float(x)
        return True
    except (ValueError, TypeError):
        return False

def _df_to_ll(df, header=True, index=True, index_label=None):
    """
    transform DataFrame or Series object into a list of lists
    :param self:
    :param header: True/False
    :param index: True/False
    :param index_label: currently unused
    :return:
    """
    if header:
        if df.columns.nlevels>1:
            if index:
                hdr = list(zip(*df.reset_index().columns.tolist()))
            else:
                hdr = list(zip(*df.columns.tolist()))
        else:
            if index:
                hdr = [df.reset_index().columns.tolist()]
            else:
                hdr = [df.columns.tolist()]
    else: hdr=[]

    if index:
        vals=df.reset_index().values.tolist()
    else:
        vals=df.values.tolist()
    return hdr + vals
