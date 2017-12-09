# -*- coding: utf-8 -*-
"""
Created on Wed Sep 29 2016

@author: Jason

Usage:
c:> python bom2dfx_embed_REF.py DXF.xls i.dxf

Excel file format:
Name(References), Quantity, Part Number, Manufacturer
저항(R2,R3,R6,R7,R11,R15,R18,R19,R22,R23,R26,R27,R30,R31)	14	CRCW06030000FKEA	Vishay(US)

for what reference is 5th column use BomwoRef-dxf.py

"""

import argparse
import dxfwrite
from dxfwrite import DXFEngine as dxf
import sys
import xlrd

#[Index, Name, Quantity, Part Number, Manufacturer]
list_col_width = [10, 100, 10, 60, 50]
ncol = len(list_col_width)
def_height_cell = 7
max_height = 36 * def_height_cell

def check_within_length(col, cell_val):
    # if len(list) is fit within cell return True
    len_total = len(cell_val)   # how many character in total
    len_max = int(list_col_width[col]/2.8) # how many character in one line
    if  len_total > len_max:
        return False
    else:
        return True
    
def eval_val(cell_val, col):
    # string is too long to fit in one line
    # split to several lines
    # nline : How many lines
    # data : new data with multiple line
    nline = 0
    #if  len_total > len_max:
    if check_within_length(col, cell_val) == False:    
        ldata = []
        if col == 1: # Name, end of line end with ','
            list_1 = cell_val.split('(')
            list_2 = list_1[1].split(',')
            line = list_1[0] + '('
            nline += 1
            for member in list_2:
                line_temp = line + member + ','
#                a = len(line_temp)
#                if (len(line_temp) > len_max):
                if check_within_length(col, line_temp) == False: 
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
                if check_within_length(col, line_temp) == False: 
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

def add_header(table, row):
        table.text_cell(row, 0, "품번", style='ctext')
        table.text_cell(row, 1, "품   명", style='ctext')
        table.text_cell(row, 2, "수량", style='ctext')
        table.text_cell(row, 3, "부 품 번 호", style='ctext')
        table.text_cell(row, 4, "제조사(제조국)", style='ctext')

def build_data(start_row, total_row, worksheet):        
    list_table = []
    list_height = []
    total_line = 0
    total_height = 0
    rows = 0
#    print "a",total_row-1-start_row
    for i in range(total_row-start_row): # last line is header
        list_table.append([])
        max_line = 1
        for col_dxf in range(ncol):
            if col_dxf == 0:
                data = str(start_row+rows+1)
            else:
                col_excel = col_dxf-1
                cell_val = worksheet.cell_value(i+start_row,col_excel)
                try:
                    float(cell_val)
                    value = str(int(cell_val)) #all number is float in excel
                except ValueError:
                    value = cell_val
                lines, data = eval_val(value, col_dxf)
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


def make_table(point, list_table, list_height):
    nrow = len(list_table)
    table = dxf.table(insert=point, nrows=nrow+1, ncols=ncol)
    # create a new styles
    ctext = table.new_cell_style('ctext', textcolor=7, textheight=3.2,
                                 halign=dxfwrite.CENTER,
                                 valign=dxfwrite.MIDDLE
                                 )
    ltext = table.new_cell_style('ltext', textcolor=7, textheight=3.2,
                                 hmargin=2,
                                 halign=dxfwrite.LEFT,
                                 valign=dxfwrite.MIDDLE
                                 )
    # modify border settings
    border = table.new_border_style(color=7)    #white boarder
    ctext.set_border_style(border)
    ltext.set_border_style(border)
                                 
    # set colum width, first column has index 0
    table.set_col_width(0, list_col_width[0])  #Index
    table.set_col_width(1, list_col_width[1])  #Name
    table.set_col_width(2, list_col_width[2])  #Quantity
    table.set_col_width(3, list_col_width[3])  #Part Number
    table.set_col_width(4, list_col_width[4])  #Manufacturer
        
    #set row height, first row has index 0
    for row in range(nrow):
        rev_nrow = nrow-1-row
        for col in range(ncol):
            if col == 0 or col == 2: #Index, Quantity
                table.text_cell(row, col, list_table[rev_nrow][col], style='ctext')
            else:
                table.text_cell(row, col, list_table[rev_nrow][col], style='ltext')            
        table.set_row_height(row, list_height[rev_nrow]*def_height_cell)
    add_header(table, nrow)
    table.set_row_height(nrow, def_height_cell+1)
    return table
    
def bom_dxf(infile, outfile):
    # Korean support
    reload(sys)
    sys.setdefaultencoding('utf-8')

    dwg = dxf.drawing(outfile) # create a drawing

    workbook = xlrd.open_workbook(infile) #open excel file
    worksheet_names = workbook.sheet_names()    # list of worksheet names
    worksheets = workbook.sheets()              # instance of worksheets

    gap = sum(list_col_width)+10    # table column position 
    j = 0
    for i, ws in enumerate(worksheets):
        total_row = ws.nrows #줄 수 가져오기
        start_row = 0
        # Print sheet name
        header = worksheet_names[i]+" ("+str(total_row)+" rows)"
        dwg.add(dxf.text(header, insert=(j*gap, 10), height=4, layer='TEXTLAYER'))
        # Table
        while (start_row < total_row):
            rows, list_table, list_height = build_data(start_row, total_row, ws)
            #print rows, list_height
            start_row += rows
            print "start:",start_row, "lenth:",len(list_table), "total:",total_row 
            table = make_table((j*gap, 0), list_table, list_height)
            j += 1
            dwg.add(table)  #Editable
            #dwg.add_anonymous_block(table, basepoint=(0, 0), insert=(i*gap, 0))    #Not editable
        
    dwg.save()
    print("drawing '%s' created.\n" % outfile)

if __name__ == '__main__':

    parser = argparse.ArgumentParser(description="Convert .xls to .dxf")
    parser.add_argument('files', nargs=2, help="infile outfile")
    # Require at least one column name
    args = parser.parse_args()			
    
    bom_dxf(args.files[0], args.files[1])