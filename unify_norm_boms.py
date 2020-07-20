# -*- coding: utf-8 -*-
"""
Created on Mon Jun 29 22:14:38 2020

@author: Liz
"""

############################################################################
##  This program has one parameter - input directory
##  The program goes through all of the .csv files in that directory
##  It puts each .csv file into a sheet of an excel file
##  It should be run on the output_bom directory which was created from the other program


import os
from glob import glob
import csv
import pandas as pd
import sys
from xlsxwriter.workbook import Workbook

def unify_norm_boms(dir_name):
    colwidths = {}
    colwidths[0] = 40
    colwidths[1] = 5
    colwidths[2] = 30
    colwidths[3] = 5
    colwidths[4] = 20
    colwidths[5] = 20
    colwidths[6] = 5
    colwidths[7] = 40
    colwidths[8] = 40
    colwidths[9] = 5
    
    output_dir_name = dir_name+"/*.csv"
    print("output_dir_name =", output_dir_name)
    workbook = Workbook(dir_name + '/all_norm_boms.xlsx', {'strings_to_numbers': True,'constant_memory': True})

    for csvfile in glob(output_dir_name):
        name = os.path.basename(csvfile).split('.')[-2]
        print("name = ", name)

        with open(csvfile, 'r') as f:
            worksheet = workbook.add_worksheet(name)
            for col_num, width in colwidths.items():
                worksheet.set_column(col_num, col_num, width)

            r = csv.reader(f)
            for row_index, row in enumerate(r):
                for col_index, data in enumerate(row):
                    worksheet.write(row_index, col_index, data)
    
    workbook.close()
    

# The following makes this program start running at normalize_all() 
# when executed as a stand-alone program.    
if __name__ == '__main__':
    unify_norm_boms(sys.argv[1])
    
