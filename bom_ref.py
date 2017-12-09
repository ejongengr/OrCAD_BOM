# -*- coding: utf-8 -*-
"""

OR-Cad BOM manupulation

@author: Jason
"""

import argparse
import pandas as pd
import numpy as np
import sys

def reference(infile, outfile):
    """
        Merge multiple line References to one line
        Input : OrCAD BOM file
        Output: BOM file to open in excel
        Usage : python bom_merge.py sensor.bom sensor_new.bom

        -> USE bom_reference.py instead of this script

        """
    newline = list()
      
    status = 'start'

    try:
        fw = open(outfile, 'w')
    except:
        print("File open error")
        sys.exit()

    # load it
    i = 0
    with open(infile) as f:
        ref_found = False;     # index of "Reference" found
        status = 'first_line'
        for line in f:
            words = line.split('\t') # TAP removed
            if not ref_found:
                k = 0
                for word in words:
                    if word == "Reference":
                        index_ref = k
                    if word == "Quantity":
                        index_qty = k
                        ref_found = True
                    k += 1 
                fw.write(line)
            else:
                # Reference index found 
                try:
                    if words[index_qty] != "":
                        if status == 'first_line':
                            status = 'merging'
                        else:    
                            fw.write("\t".join(newline)) # joint with Tab
                            # check quantity
                            refs = newline[index_ref].split(',')
                            if len(refs) != int(newline[index_qty]):
                                print('Quantity mismatch :' + str(newline[index_qty]) + str(refs))
                        newline = words
                    else:
                        refs = words[index_ref].rstrip()
                        newline[index_ref] += refs
                except:
                    fw.write(line)
                        
    fw.write("\t".join(newline)) #last one
                             
    # save it
    fw.close()
    
def concat():
    file_cnt = len(args.files)
    if  file_cnt < 3:
        msg = "argument requires more than 2 files"
        raise argparse.ArgumentTypeError(msg)
        sys.exit()

    print(args.files)

    list_df = []
        
    # stack and sort 
    for i in range(file_cnt-1):
        print(i)
        print(args.files[i])
        dft = pd.read_csv(args.files[i], sep='\t', skiprows = 10)
        dft = dft.drop(0)
        dft["file"] = args.files[i]
        list_df.append(dft)
    dfa = pd.concat(list_df).sort_values(by="PN")
    dfa.to_csv(args.files[-1], sep='\t', index=False)
    
def merge():
    fd    