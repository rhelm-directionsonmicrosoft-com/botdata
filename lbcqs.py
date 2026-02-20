# lbcqs
import pandas as pd
import os
import os.path
import time

inpath = r"C:\Users\Rob\script\botdata\lbc"
outpath  = r"C:\Users\Rob\script\botdata\lbcout.csv"

def combine_excel_files(inpath=inpath, outpath=outpath):
    result = concat_files(inpath, pd.DataFrame())
    print(f'concat_files result {result.info()}')
    with open(outpath, mode='w', encoding='utf-8') as outf:
        result.to_csv(outpath, index=False)
    outf.close()
    return(result)


def concat_files(inpath, result):
    print(f'concat_files start {result.info()}')
    for filename in os.scandir(inpath):
        print(f'concat_files filename.path = {filename.path}')
        if is_excel(filename):
            print(f'is excel file')
            with pd.ExcelFile(filename) as xlfile:
                result = concat_worksheets(xlfile, filename.path, result)
            xlfile.close()
        else:
            print('is not excel file')
    return(result)

def concat_worksheets(xlfile, path, result):
    sheet_names = xlfile.sheet_names
    for n in sheet_names:
        print(f'concat_worksheet {n}')
        new_sheet = pd.DataFrame = pd.read_excel(xlfile, n)
        new_sheet.name = n
        new_sheet.insert(len(new_sheet.columns), "source_path", str(path))
        new_sheet.insert(len(new_sheet.columns), "source_sheet", n)
        result = pd.concat([result, new_sheet], ignore_index=True)
    print(f'concat_worksheets {result.info()}')
    return(result)
    
def is_excel(filename):
    path = filename.path
    if os.path.isfile(path):
        (base,extension) = os.path.splitext(path)
        if extension == '.xlsx':
            print('is_excel: True')
            return(True)
        else:
            print('is_excel: False')
            return(False)
    else:
        print('is_excel: False')
        return(False)

            
        
