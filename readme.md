OrCAD BOM
=====

These are python scripts with OrCAD BOM files

Highlights
-------

* convert bom to excel
* merge serveral bom to one excel
* When merging, can use others.xls
* When merging, can use BOM database excel file

System Requirements
-------
Python 2.7

Usage
-------

* conversion
```
C:>python bom_reference.py 1.bom 1.xls
BOM Reference merged file " 1.xls " has been created
```
* merging
```
C:>python bom_mergy.py 1.bom 2.bom 3.bom all.xls

BOM file for Multiple board " all.xls " has been created
```
* merging with DB
```
C:>python bom_mergy.py -db bom_db.xls -inc others_bom.xls 1.bom 2.bom 3.bom all_db.xls

BOM file for Multiple board " all_db.xls " has been created
```
* merging with DB and others
```
C:>python bom_mergy.py -db bom_db.xls -inc others_bom.xls 1.bom 2.bom 3.bom all_db_others.xls

BOM file for Multiple board " all_db_others.xls " has been created
```

Notes
-------
* bom_reference.py and bom_reference_list.py do the same job

License
-------
Free to use
