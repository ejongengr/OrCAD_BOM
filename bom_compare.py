# -*- coding: utf-8 -*-
import argparse
import numpy as np
import pandas as pd
import re

import bom_reference as br

def bom_c(a, b):
    """
     a : source string
     b : target string
    """
    r = re.compile('C\d\d\d') # 예로 C234
    if any(x in b for x in a):
        print 
#contained = [x for x in d if x in paid[j]]    
#firstone = next((x for x in d if x in paid[j]), None)

def to_each(dfi):
    """
        input BOM: each part number has each line 
        output BOM: each reference has each line
    """
    dfo = None
    for i in range(len(dfi)):
        ref = dfi['Reference'].iloc[i]
        ref_lst = ref.split(',')
        ref_len = len(ref_lst)    
        ser = dfi.iloc[i,:].copy()
        dft = pd.DataFrame([ser]*ref_len)
        dft['Reference'] = ref_lst
        if dfo is not None:
            dfo = pd.concat([dfo, dft], ignore_index=True)
        else:
            dfo = dft.copy()
    return dfo
    
def comp_ref_changed(df_old_each, df_new_each):
    """
        check chagned, deleted, added reference
    """
    columns=['ref','old','new','old_mount', 'new_mount', 'remark']
    df = pd.DataFrame([], columns=columns)
    count_changed = 0
    count_deleted = 0
    count_added = 0
    count_mount = 0
    drop_new_lst = []
    df_old_each = df_old_each.fillna('')
    df_new_each = df_new_each.fillna('')
    count_total_new = 0
    for i, [ref_old, part_old, mount_old] in enumerate(zip(df_old_each['Reference'], df_old_each['Part'], df_old_each['Mount'])):
        found = False
        for j, [ref_new, part_new, mount_new] in enumerate(zip(df_new_each['Reference'], df_new_each['Part'], df_new_each['Mount'])):
            if ref_old == ref_new:
                found = True
                drop_new_lst.append(j)
                if part_old != part_new: # changed
                    count_changed += 1
                    lst = [ref_old, part_old, part_new, mount_old, mount_new,'수정']
                    df = df.append(pd.DataFrame([lst], columns=columns))
                else: # not changed                    
                    if mount_old != mount_new: # mount option changed
                        count_mount += 1
                        if 'DNP' == mount_new:
                            lst = [ref_old, part_old, part_new, mount_old, mount_new,'조립안함']
                        elif 'DNP' == mount_old:
                            lst = [ref_old, part_old, part_new, mount_old, mount_new,'조립함']
                        else:    
                            lst = [ref_old, part_old, part_new, mount_old, mount_new,'알수없음']
                        df = df.append(pd.DataFrame([lst], columns=columns))
                break;
        if found == False: #deleted
            count_deleted += 1
            lst = [ref_old, part_old, '', mount_old, '','삭제']
            df = df.append(pd.DataFrame([lst], columns=columns))
    drop_new_lst = list(set(drop_new_lst))        
    df_new_drop = df_new_each.drop(drop_new_lst)    
    for ref, part, mount in zip(df_new_drop['Reference'], df_new_drop['Part'],  df_new_drop['Mount']):
        count_added += 1
        lst = [ref, '', part, '', mount,'추가']
        df = df.append(pd.DataFrame([lst], columns=columns))
    count_lst = [count_changed, count_deleted, count_added, count_mount]
    df = df.fillna('')
    df.reset_index(inplace=True, drop=True)
    df.index = df.index+1
    return  count_lst, df    

def comp_part_changed(df_old, df_new):
    columns=['part','old_qty','new_qty', 'mount', 'old_ref', 'new_ref', 'remark']
    df = pd.DataFrame([], columns=columns)
    df_old_p = df_old[df_old['Mount'] != 'DNP']
    df_old_p.reset_index(inplace=True)
    df_old_dnp = df_old[df_old['Mount'] == 'DNP']
    df_old_dnp.reset_index(inplace=True)
    df_new_p = df_new[df_new['Mount'] != 'DNP']
    df_new_p.reset_index(inplace=True)
    df_new_dnp = df_new[df_new['Mount'] == 'DNP']
    df_new_dnp.reset_index(inplace=True)
    # changed
    for df_old_x, df_new_x, mount in zip([df_old_p, df_old_dnp], [df_new_p, df_new_dnp], ['Mount', 'DNP']):
        drop_new_lst = []
        for i, [part_old, qty_old, ref_old] in enumerate(zip(df_old_x['Part'], df_old_x['Quantity'], df_old_x['Reference'])):
            found = False
            for j, [part_new, qty_new, ref_new] in enumerate(zip(df_new_x['Part'], df_new_x['Quantity'], df_new_x['Reference'])):
                if part_old == part_new:
                    found = True
                    drop_new_lst.append(j)
                    if qty_old != qty_new: # qty changed
                        lst = [part_old, qty_old, qty_new, mount, ref_old, ref_new, 'Qty 수정']
                        df = df.append(pd.DataFrame([lst], columns=columns))
                    elif ref_old != ref_new: # ref changed
                        lst = [part_old, qty_old, qty_new, mount, ref_old, ref_new, 'Ref 수정']
                        df = df.append(pd.DataFrame([lst], columns=columns))                    
                    break
            if found == False: #deleted
                lst = [part_old, qty_old, 0, mount, ref_old, '', '삭제']
                df = df.append(pd.DataFrame([lst], columns=columns))
        drop_new_lst = list(set(drop_new_lst))
        df_new_drop = df_new_x.drop(drop_new_lst)    
        for part, qty, ref in zip(df_new_drop['Part'], df_new_drop['Quantity'], df_new_drop['Reference']):
            lst = [part, 0, qty, mount, '', ref, '추가']
            df = df.append(pd.DataFrame([lst], columns=columns))
        df = df.fillna('')
        df['old_qty'] = df['old_qty'].astype('int')
        df['new_qty'] = df['new_qty'].astype('int')
    df.reset_index(drop=True, inplace=True)
    df.index = df.index+1
    return df                

def com_add_mount(df):
    if 'Mount' not in df.columns:
        df['Mount'] = ''
    return df
    
# call main
if __name__ == '__main__':
    # create parser
    parser = argparse.ArgumentParser(description="compare tow BOMs")
    # add expected arguments
    parser.add_argument('files', nargs=2, help="old_file new_file")
    parser.add_argument('-o', '--o', default=None, help="output file name")
    # parse args
    args = parser.parse_args()
    
    df_old = br.reference(args.files[0], None, None, 'Part')
    # change nc to DNP
    df_old = com_add_mount(df_old)
    df_old_each = to_each(df_old)

    df_new = br.reference(args.files[1], None, None, 'Part')
    # change nc to DNP
    df_new = com_add_mount(df_new)
    df_new_each = to_each(df_new)

    # reference changed
    count_lst, dfr = comp_ref_changed(df_old_each, df_new_each)

    # part count changed
    dfp = comp_part_changed(df_old, df_new)

    print ('Old file = %s, New file = %s' % (args.files[0], args.files[1]))
    print ('Total old BOM part count = %d' % len(df_old_each))
    print ('Total new BOM part count = %d' % len(df_new_each))
    print ('Changed part count = %d' % count_lst[0])
    print ('Deleted part count = %d' % count_lst[1])
    print ('Added part count = %d' % count_lst[2])
    print ('Mount option changed = %d' % count_lst[3])
    if args.o == None:
        fn = 'bom_diff.xls'
    else:
        fn = args.o
        
    writer = pd.ExcelWriter(fn)   
    dfr.to_excel(writer, 'Reference', index_label='index')
    dfp.to_excel(writer, 'Part count', index_label='index')
    writer.save()
    print ('%s saved' % fn)
    
    