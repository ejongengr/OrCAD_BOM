# -*- coding: utf-8 -*-
"""
Created on Wed Sep 29 2016

@author: Jason

Usage:
c:> python bom2dxf.py DXFwoRef.xls name.dxf

Excel file format:

품번	품   명	수량	부 품 번 호	제조사(제조국)	부품기호(실장위치)
10  25	    10  50	    45	        100
c   l       c   l       l           l   
1	직접회로	23	TPS73133DBVT	TI(US)	U3
2	직접회로	1	TLV1548QDBREP	TI(US)	U4
3	볼륨	7	SW45009	Electo-Mech(US)	POT1,POT2,POT3,POT4,POT5,POT6,POT7,POT1,POT2,POT3,POT4,POT5,POT6,POT7,POT1,POT2,POT3,POT4,POT5,POT6,POT7, POT8
4	커넥터	33	SMPL20S0N0LB	Positronic(US)	CON1

first line: header
second line: column width
third line: text alignment, center, left or right

#[Index, Name, Quantity, Part Number, Manufacturer]
list_col_width = [10, 25, 10, 50, 45]
list_col_name = ["품번", "품   명", "수량", "부 품 번 호", "제조사(제조국)"]

#[Index, Reference]
list_ref_width = [10, 100]
list_ref_name = ["품번", "부품기호(실장위치)"]
"""

import argparse
import dxfwrite
from dxfwrite import DXFEngine as dxf
import numpy as np
import pandas as pd
import sys

def_height_cell = 8
max_height = 252

def eval_val(cell_val, len_max, iscomma):
    # string is too long to fit in one line
    # split to several lines
    # nline : How many lines
    # data : new data with multiple line
    nline = 0
    #if  len_total > len_max:
    len_cell = len(cell_val)
    if len_cell > len_max: 
        ldata = []
        if iscomma: # Name, end of line end with ','
            list_ref = cell_val.split(',')
            line = ""
            nline += 1
            for member in list_ref:
                line_temp = line + member + ','
                if len(line_temp) > len_max:
                    ldata.append(line)
                    ldata.append('\n')                
                    line = member + ','
                    nline += 1
                else:    
                    line = line + member + ','                        
            ldata.append(line[:-1]) # remove last comma (,)
        else:
            line = ""
            for member in cell_val:
                line_temp = line + member
                if len(line_temp) > len_max:
                    ldata.append(line)
                    ldata.append('\n')                
                    line = member
                    nline += 1
                else:    
                    line = line + member
            ldata.append(line) # remove last comma (,)
        data = ""            
        for member in ldata:
            data = data + member
    else:
        data = cell_val
    return nline, data

def build_data(lst_colwidth, start_row, total_row, df, comma_idx=None, scale_factor=2.8):
    #
    # comma_idx : column index of reference
    # auto_index : auto generate index
    #
    list_table = []
    list_height = []
    total_line = 0
    total_height = 0
    rows = 0
    for i in range(total_row-start_row): # last line is header
        list_table.append([])
        max_line = 1
        for col in range(len(lst_colwidth)):        
            cell_val = df.iloc[i+start_row,col]
            
            try:
                float(cell_val)
                if np.isnan(cell_val):     # cell is empty 
                    cell_val = " "
                cell_val = str(int(cell_val)) #all number is float in excel
            except ValueError:
                cell_val = cell_val # text
                        
            len_max = int(lst_colwidth[col]/scale_factor)
            if col == comma_idx:
                lines, data = eval_val(cell_val, len_max, True);
            else:
                lines, data = eval_val(cell_val, len_max, False);
            
            if lines > max_line:
                max_line = lines   # keep maximum line
            list_table[i].append(data)
        list_height.append(max_line)
        rows += 1
        total_line += max_line
        total_height = total_line * def_height_cell
        if total_height > max_height or start_row+rows > total_row:
            break
    return rows, list_table, list_height   

def make_one_table(point, list_table, list_height, lst_colwidth, lst_colalign, reverse):
    nrow = len(list_table)
    ncol = len(lst_colwidth)
    table = dxf.table(insert=point, nrows=nrow+1, ncols=ncol) # nrow+1 : count header
    ## 7=white, 0=BYBLOCk, 256=BYLAYER
    
    # create a center alignment styles 
    ctext = table.new_cell_style('ctext', textcolor=7, textheight=3.2,
                                 halign=dxfwrite.CENTER,
                                 valign=dxfwrite.MIDDLE
                                 )
    # create a right alignment styles 
    rtext = table.new_cell_style('ltext', textcolor=7, textheight=3.2,
                                 hmargin=2,
                                 halign=dxfwrite.RIGHT,
                                 valign=dxfwrite.MIDDLE
                                 )
    # create a left alignment styles 
    ltext = table.new_cell_style('ltext', textcolor=7, textheight=3.2,
                                 hmargin=2,
                                 #linespacing=0.5,
                                 halign=dxfwrite.LEFT,
                                 valign=dxfwrite.MIDDLE
                                 )
    # modify border settings
    border = table.new_border_style(color=7)    #white boarder = 7
    ctext.set_border_style(border)
    ltext.set_border_style(border)
                                 
    # set colum width, first column is index 0
    for i, width in enumerate(lst_colwidth):
        table.set_col_width(i, width)
        
    #set row height, first row is index 0
    for row in range(nrow):
        if reverse:
            newrow = nrow-1-row    # last line is header
            tabel_row = row
        else:
            newrow = row
            tabel_row = row+1   # first line is header
            
        for col in range(ncol):  # text alignment
            if lst_colalign[col] == "C" or lst_colalign[col] == "c" : # center
                table.text_cell(tabel_row, col, list_table[newrow][col], style='ctext')
            elif lst_colalign[col] == "R" or lst_colalign[col] == "r": # right
                table.text_cell(tabel_row, col, list_table[newrow][col], style='rtext')
            else: # left 
                table.text_cell(tabel_row, col, list_table[newrow][col], style='ltext')
            
        table.set_row_height(tabel_row, list_height[newrow]*def_height_cell)
    return table

def make_tables_horizontal(dwg, dfs, pos_x, pos_y, comma_idx=None, reverse=True, scale_factor=2.8):
    # dwg: drwaings
    # dfs: datafram dictionary
    # lst_colwidth: list fo column width
    # lst_col_name: list of column name
    # pos_x: x position of tabel
    # pos_y: y posion of table
    # comma_idx: column index where comma sperable reference is
    # reverse: table descend order
    # Usage:
    #   reload(sys)
    #   sys.setdefaultencoding('utf-8') # support Korea
    #   dwg = dxf.drawing(outfile) # create a drawing
    #   dfs = pd.read_excel(infile, sheetname=None)
    #   dwg = dwg = make_tables_horizontal(dwg, dfs, 0, 10, None, True)
    #   dwg.save()

    j = 0
    for df_name in dfs:
        df = dfs[df_name]
        lst_colname = df.columns.values  # retrieve header name
        lst_colwidth = list(map(int, df.loc[0,:])) # first row is column width        
        lst_colalign = list(df.loc[1,:]) # first row is column alignment
        df = df.drop(df.index[[0,1]]) # drop first 2 row which is column width and alignment
        
        gap = sum(lst_colwidth)+10    # table column position 
        total_row = len(df.index) #줄 수 가져오기
        start_row = 0
        # Print sheet name
        header = df_name+" ("+str(total_row)+" rows)"
        print header
        dwg.add(dxf.text(header, insert=(j*gap, pos_y), height=4, layer='TEXTLAYER'))
        # Table
        while (start_row < total_row):            
            rows, list_table, list_height = build_data(lst_colwidth, \
                                            start_row, total_row, df, comma_idx, scale_factor);
            start_row += rows
            print "start:",start_row, "lenth:",len(list_table), "total:",total_row
            table = make_one_table((j*gap, pos_y-10), list_table, list_height, \
                                    lst_colwidth, lst_colalign, reverse)
            for i, name in enumerate(lst_colname):  # header
                if reverse:
                    table.text_cell(len(list_table), i, lst_colname[i], style='ctext')
                    table.set_row_height(len(list_table), def_height_cell+1)
                else:  # first line
                    table.text_cell(0, i, lst_colname[i], style='ctext') 
                    table.set_row_height(0, def_height_cell+1)
            j += 1
            dwg.add(table)  #Editable
            #dwg.add_anonymous_block(table, basepoint=(0, 0), insert=(i*gap, 0))    #Not editable
    
    return dwg
    
def bom_dxf(infile, outfile, ref_column_index, reverse, scale_factor):
    # Korean support
    reload(sys)
    sys.setdefaultencoding('utf-8')    
    dwg = dxf.drawing(outfile) # create a drawing
    dfs = pd.read_excel(infile, sheetname=None)
    
    # bom table
    dwg = make_tables_horizontal(dwg, dfs, 0, 10, ref_column_index, reverse, scale_factor)

    dwg.save()
    print("drawing '%s' created.\n" % outfile)

if __name__ == '__main__':

    parser = argparse.ArgumentParser(description="Convert .xls to .dxf")
    parser.add_argument('files', nargs=2, help="infile outfile")
    parser.add_argument('-ref', default=None, help="comma seperable column(ex. Reference) index")
    parser.add_argument('-reverse', action='store_true', help="comma seperable column(ex. Reference) index")
    parser.add_argument('-scale', default=2.8, help="how many point in one character width")    # 2.0
    
    # Require at least one column name
    args = parser.parse_args()			
    
    bom_dxf(args.files[0], args.files[1], args.ref, args.reverse, args.scale)