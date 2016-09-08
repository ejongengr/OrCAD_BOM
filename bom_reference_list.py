# -*- coding: utf-8 -*-
"""
Created on Wed Dec  9 17:29:56 2015

Merge multiple lines of References to one line

Input : OrCAD BOM file
Output: BOM file to open in excel
Usage : python bom_reference_list.py 1.bom 1.xls

@author: Jason
"""
import argparse
import sys
import xlwt

def reference(infile, outfile):
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
                                print 'Quantity mismatch :' + str(newline[index_qty]) + str(refs)
                        newline = words
                    else:
                        refs = words[index_ref].rstrip()
                        newline[index_ref] += refs
                except:
                    fw.write(line)
                        
    fw.write("\t".join(newline)) #last one                             
    # save it
    fw.close()
    
    ##################################################################
    # Write to excel
    workbook = xlwt.Workbook(encoding='utf-8')
    #workbook.default_style.font.height = 20*11
    worksheet = workbook.add_sheet(u'Sheet1') 
    #font_style = xlwt.easyxf('font:height 280;')
    #worksheet.row(0).set_style(font_size)    
    skiplst = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13]
    with open(outfile) as fe:
        rdrow = 0   # read row
        wrrow = 0   # write row
        for line in fe:
            words = line.split('\t') # TAP removed
            col = 0
            if rdrow not in skiplst:
                for word in words:
                    worksheet.write(wrrow, col, word)
                    col += 1                    
                wrrow += 1
            rdrow += 1    
    fs = outfile.split('.')
    fn = fs[0]+".xls"
    workbook.save(fn)
    
# call main
if __name__ == '__main__':
    # create parser
    parser = argparse.ArgumentParser(description="Merge reference using binary")
    # add expected arguments
    parser.add_argument('files', nargs=2, help="infile outfile")
    # parse args
    args = parser.parse_args()
    reference(args.files[0], args.files[1])
