# -*- coding: utf-8 -*-
"""
Created on Wed Dec  9 15:19:33 2015

Convert excel save as text format to OrCAD property input file

Input : Excel save as text format
Output: OrCAD import paramter file
Usage : python property_convert sensor.exp sensor_new.exp


@author: Jason
"""
import sys
   
if len(sys.argv) != 3:
    print("Convert excel save as text format to OrCAD property input file")    
    print("Usage:python property_convert input_file output_file")
    sys.exit()

infile = sys.argv[1]    
outfile = sys.argv[2]

# load it
f = open(infile, 'rb')
data = f.read()
f.close()

newdata = bytearray()

# do something with data
newdata.append('\x22') # add " at first line
for d in data:
    if d != '\x22': #delete ""
        if d == '\x09':  # if tap and add ""
            newdata.append('\x22')
            newdata.append('\x09')
            newdata.append('\x22')
        elif d == '\x0D':   # if carrage-return and add ""
            newdata.append('\x22')            
            newdata.append('\x0D')            
            newdata.append('\x0A')            
            newdata.append('\x22')            
        elif d == '\x0A':
            pass                 
        else:
            newdata.append(d)
                   
# save it
f = open(outfile, 'wb')
f.write(newdata)
f.close()