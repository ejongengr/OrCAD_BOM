# -*- coding: utf-8 -*-
"""
Created on Sun Feb 14 11:24:10 2016

Usage : python bom_merge.py 1.bom 2.bom 3.bom output.xls
        // Include other bom which is excel file
		python bom_merge.py 1.bom 2.bom 3.bom -inc other.xls output.xls
        // Use database excel file
		python bom_merge.py 1.bom 2.bom 3.bom -inc other.xls -db db.xls output.xls
        // Use on key for merging, default on key is "Part"
		python bom_merge.py 1.bom 2.bom 3.bom -inc other.xls -db db.xls output.xls -on Value
@author: Jason Lee
"""
import argparse
import numpy as np
import pandas as pd
from pandas import ExcelWriter
import sys

from bom_reference import reference

def concat(list_dfs, lst_filename, incfilename, onkey):
    for i in range(len(list_dfs)):
        list_dfs[i]["file"] = lst_filename[i]
    if incfilename != None:
        list_dfs[len(list_dfs)-1]["file"] = incfilename
        
    dfc = pd.concat(list_dfs).sort_values(by=onkey)
    #save concaternated file
    fs = lst_filename[-1].split('.')
    fn = fs[0] + "_concat"     # file name - *_concat.*    
    return dfc, fn

def merge(dfa, outfile, onkey):
    sum = dfa.groupby([onkey], as_index=False).agg(np.sum) #Nou use data as index, Sum quantity
    dfa = dfa.drop(['Reference', 'Quantity', 'file'], axis = 1) #group by Part
    dfa = dfa.drop_duplicates([onkey]) # remove Part duplication
    
    dfm = sum.merge(dfa, on=onkey) # merge quantity with origin
    dfm = dfm.sort_values(["Item", onkey])
    dfm.index = list(range(1, len(dfm)+1))    # reset and sort index again
    
    fs = outfile.split('.')
    fn = fs[0] + "_All"
    return dfm, fn
    
# call main
if __name__ == '__main__':
    # create parser
    parser = argparse.ArgumentParser(description="Stack multiple BOM to one")
    # add expected arguments
    parser.add_argument('files', nargs='+', help="multiple input files and one output file")
    parser.add_argument('-inc', nargs=1, help="include xls file")
    parser.add_argument('-db', nargs=1, help="use db files")
    parser.add_argument('-on', nargs='?', default='Part', help="use on key for merge")
    # parse args
    args = parser.parse_args()
    
    onkey = args.on
    list_dfs = []
    file_cnt = len(args.files)
    if not args.db:
        dbfile = None
    else:
        dbfile = args.db[0]        
    for i in range(file_cnt-1):
        list_dfs.append(reference(args.files[i], None, dbfile, onkey))
    
    if args.inc:
        dfinc = pd.read_excel(args.inc[0])
        dfinc.index = list(range(1, len(dfinc)+1))    # reset and sort index again
        list_dfs.append(dfinc)
    
    if args.inc:
        incfilename = args.inc[0]
    else:
        incfilename = None
    dfc, fnc = concat(list_dfs, args.files, incfilename, onkey)
    dfm, fnm = merge(dfc, args.files[-1], onkey)

    fs = args.files[-1].split('.')
    fn = fs[0]+".xls"
    with ExcelWriter(fn) as writer:
        for i in range(file_cnt-1):
            list_dfs[i].to_excel(writer, sheet_name=args.files[i])
        if args.inc:
            list_dfs[-1].to_excel(writer, sheet_name=args.inc[0])
        dfc.to_excel(writer, sheet_name=fnc)
        dfm.to_excel(writer, sheet_name=fnm)
        
    print("\nBOM file for Multiple board \"", fn, "\" has been created")
