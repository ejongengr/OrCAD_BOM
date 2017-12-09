# -*- coding: utf-8 -*-
"""
Created on Wed Dec  9 17:29:56 2015

Merge multiple line References to one line

Input : OrCAD BOM file
Output: BOM file to open in excel
Usage : python bom_reference.py sensor.bom sensor.xls
        // Use database excel file
        python bom_reference.py sensor.bom sensor.xls -db db.xls
        // Use on key for merging, default on key is "Part"
		python bom_reference.py sensor.bom sensor.xls -db db.xls -on Value
        
        import bom_reference as br
        dfn = br.reference('icp_front.BOM', None, None, 'Part')
        
@author: Jason
"""
import argparse
import pandas as pd

def reference(infile, outfile, dbfile, onkey):
      
    # load it
    df = pd.read_csv(infile, sep='\t', skiprows = 10) #, na_filter=False)
    df = df.drop(0)
    
    dfNull = df.isnull()
    for idx in df.index:
        n = idx - 1
        if n == 0:
            # first line
            dfNew = pd.DataFrame([list(df.iloc[0])], columns=df.columns)
            #	Item	Quantity	Mount Reference	Value	Part	Manufacturer	Description	Spec
            #	Bead	1	DNP BD1	HB-4M3216A-121JT	HB-4M3216A-121JT	Ceratech	150mA

            dfTemp = df.iloc[0].copy()
            strTemp = df.iloc[0]["Reference"]
            dfNew = dfNew.drop(0) #delete first line again -> empty dataframe
        else:  
            if dfNull.iloc[n]["Item"] == True: # Reference to add
                strTemp += df.iloc[n]["Reference"]
            else: # Add done
                dfNew = dfNew.append(dfTemp.replace(dfTemp["Reference"], strTemp))
                dfTemp = df.iloc[n].copy()
                strTemp = df.iloc[n]["Reference"]
    dfNew = dfNew.append(dfTemp.replace(dfTemp["Reference"], strTemp)) #add last
    
    if dbfile != None:
        dfb = pd.read_excel(dbfile)
        # Left columns, *.BOM text file
        if 'Mount' in dfNew.columns:
            dfNew = dfNew[[onkey, "Reference", "Quantity", "Mount"]]
        else:    
            dfNew = dfNew[[onkey, "Reference", "Quantity"]]
        # Right columns refer to database, BOM.xls
        dfb = dfb[[onkey, "Item", "Manufacturer", "Description", "Spec", "Package"]]
        dfNew = dfNew.merge(dfb, how="left", on=onkey, )
        if 'Mount' in dfNew.columns:
            dflist = ["Item", "Quantity", "Reference", "Mount", onkey, "Manufacturer",
                            "Description", "Spec", "Package"]
        else:
            dflist = ["Item", "Quantity", "Reference", onkey, "Manufacturer",
                            "Description", "Spec", "Package"]
        dfNew = dfNew[dflist]
    
    dfNew["DNP"] = dfNew['Part'].astype(str).str.contains('/NC')
        
#    if onkey in dfNew:
    dfNew = dfNew.sort_values(["DNP", "Item", onkey])
#    else:
#        dfNew = dfNew.sort_values(["DNP", "Item"])
    
    dfNew.index = list(range(1, len(dfNew)+1))    # reset and sort index again
    if outfile != None:
        fs = outfile.split('.')
        fn = fs[0]+".xls"
        dfNew.to_excel(fn)
        print("BOM Reference merged file \"", fn, "\" has been created")
    return dfNew
    
# call main
if __name__ == '__main__':
    # create parser
    parser = argparse.ArgumentParser(description="Merge reference using pandas")
    # add expected arguments
    parser.add_argument('files', nargs=2, help="infile outfile")
    parser.add_argument('-db', nargs=1, help="use db files")
    parser.add_argument('-on', nargs=1, default="Part", help="use on key for merge")
    # parse args
    args = parser.parse_args()
    if not args.db:
        dbfile = None
    else:
        dbfile = args.db[0]        
    reference(args.files[0], args.files[1], dbfile, args.on)
