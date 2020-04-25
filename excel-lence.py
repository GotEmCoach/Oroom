#!/usr/bin/env python3

import openpyxl
from pathlib import Path
import os
import time
import shutil

def master_eat_children(mastername, children, m, backupdir, actualdir, cleanfile):
    if actualdir.exists() == True:
        with open(actualdir, 'rb') as f:
            with open(backupdir, 'wb+') as f2:
                f2.write(f.read())
        os.remove(actualdir)
        shutil.copyfile(cleanfile, actualdir)
    for i in children:
        masterwb = openpyxl.load_workbook(actualdir)
        childwb = openpyxl.load_workbook(i+'.xlsx') 
        masterws = masterwb.worksheets[m] 
        eachchild = len(childwb.worksheets)
        for i in range(eachchild):
            childws = childwb.worksheets[i]
            masterwb.create_sheet(title=childws.title)
            mr = childws.max_row 
            mc = childws.max_column 
            for i in range (1, mr + 1): 
                for j in range (1, mc + 1): 
                    c = childws.cell(row = i, column = j) 
                    masterws.cell(row = i, column = j).value = c.value
            masterwb.save(str(mastername)) 
            m = m + 1

def main():
    mastername = input('Please enter the Name of the Master File: \n') + '.xlsx'
    children = input('Please enter the names of the files you want to merge, Seperate them by ,: \n')
    children = children.split(',')
    m = 0
    cleanfile = Path('c:\\Users\\mclar\\oroom\\Oroom\\clean.xlsx')
    actualdir = Path('c:\\Users\\mclar\\oroom\\Oroom\\' + mastername)
    backupdir = Path('c:\\Users\\mclar\\oroom\\Oroom\\backup\\'+ time.ctime().replace(":","_") + mastername)
    master_eat_children(mastername, children, m, backupdir, actualdir, cleanfile)


if __name__ == "__main__":
    main()


