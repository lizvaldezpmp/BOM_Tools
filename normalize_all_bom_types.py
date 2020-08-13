# -*- coding: utf-8 -*-
"""
Created on Sun Jun 28 15:10:28 2020

@author: Liz
"""


#########################################################################################################################
### This program has two input parameters
###   -  input directory name where the input BOMs are to be placed
###   -  output directory - This is where the csv version of the unified BOMs will go.  This will include the sort key (hash).
#########################################################################################################################
import pandas as pd
import csv
import sys
import os
from pathlib import Path
#from pathlib import path
#from IPython.external.path import path as path

#########################################################################################################################
### I would like to not have these as global variables.
#########################################################################################################################
global row_num_found, xp_pn_found,  mfg_pn_found, quantity_found,description_found, level_found, ref_des_found
global unit_of_measure_found,  manufacturer_found, comp_text_found, rev_level_found
global bom_tab_name, first_tab_of_input_bom, bom_type
global xp_sort_key_col_name, mfg_sort_key_col_name, file_col_name, final_column_order
global row_num, xp_pn, mfg_pn, quantity, description, level_val, ref_des
global unit_of_measure, manufacturer, comp_text, rev_level

#########################################################################################################################
### I was going to use these as indices in to a tuple that has the unified name of the column header and whether or not it was found.
### But the strategy didn't pan out.  I am commenting these out for not just in case I end up using them.  
#########################################################################################################################
## global row_num_index, xp_pn_index, mfg_pn_index, quantity_index, description_index
## global level_val_index, ref_des_index, unit_of_measure_index, manufacturer_index, comp_text_index, rev_level_index

#########################################################################################################################
### Initiatlize global variables. 
### I don't really know how global variables work, but I find that if I call them out as global everywhere, the program works. :)
#########################################################################################################################

def initialize_data():
    global row_num, xp_pn, mfg_pn, quantity, description, level_val, ref_des
    global unit_of_measure, manufacturer, comp_text, rev_level
    global xp_sort_key_col_name, mfg_sort_key_col_name, file_col_name, final_column_order
    global cost_bom_tab_names
    global row_num_found, xp_pn_found,  mfg_pn_found, quantity_found,description_found, level_found, ref_des_found
    global unit_of_measure_found,  manufacturer_found, comp_text_found, rev_level_found
    

###    This is the possible name of the tab to find in the 001 and VN WIP BOMs.  
###    If other (older) 001 and VN BOMs are found with other tab names, just add them to this list.

    cost_bom_tab_names = ["Cost Bom F-140-F", "Cost BOM F-140", "Cost Bom F-140", "Costed BOM", "Cost BOM"]

### These are names of columns that are going to be added in addition to the columns that are in the BOMs already
   
    xp_sort_key_col_name = "XP Sort Key"
    mfg_sort_key_col_name = "Mfg Sort Key"
    file_col_name = "File"

### These are the names of the columns that could be in any of the original BOM files.  
### If additional possible column names are found to exist in the original BOMs, just add that name to the list.  It must be in lower case in the list.
### The first value of each list will be used as the unified column name.
### Ideally, these would not be global variables.  
### Ultimately, it would be great to have these in a configuration file.

    row_num = ["No.","no.", "item number", "row"]
    xp_pn = ["XP PART NUMBERS", "xp part numbers", "part number", "internalcomponentitemnumber"]
    mfg_pn = ["MFR PART NUMBER", "mfr part number", "component number","part #"]
    quantity = ["QTY", "qty", "requiredquantity", "comp. qty (bun)"]
    description = ["DATABASE DESCRIPTION", "database description", "description", "componentitemdescription", "object description"]
    level_val = ["BOM LEVEL", "bom level", "level", "explosion level"]
    ref_des = ["REF DES", "ref des", "bom ref. des.", "referencedesignatorsum"]
    unit_of_measure= ["U/M", "u/m", "unit of measure", "componentitemum", "component uom"]
    manufacturer= ["Manufacturer", "manufacturer", "mfr"]
    comp_text = ["Text", 'componenttext', 'text']
    rev_level = ["Rev level", "revision level"]

### Ideally these would not be global variables.
    row_num_found = False
    xp_pn_found = False
    mfg_pn_found = False
    quantity_found = False
    description_found = False
    level_found = False
    ref_des_found = False
    unit_of_measure_found = False
    manufacturer_found = False
    comp_text_found = False
    rev_level_found = False

######################################################################################################################### 
### This is a stub so far.  It is to check to see that all required BOM fields are populated.
#########################################################################################################################

def check_bom_fields (input_file, sheet_name, bom_type, df, col_order):
#    print("len of df is", len(df))
#    print("final column order is", col_order)
    for row_index, row in df.iterrows():
        if df.loc[row_index, xp_pn[0]] != "":
#            print("row ", row_index, "level ", df.loc[row_index,level_val[0]],"part number is", df.loc[row_index, xp_pn[0]])

            if str(df.loc[row_index,level_val[0]]) == "":
                print (input_file, ":", sheet_name, "-- Bom level missing")
                
            if str(df.loc[row_index,quantity[0]]) == "":
                print (input_file, ":", sheet_name, "-- Quantity missing")
                
            if str(df.loc[row_index,quantity[0]]) == "":
                print (input_file, ":", sheet_name, "-- Quantity missing")

            if "wire" in str(df.loc[row_index,description[0]]):
                if str(df.loc[row_index,unit_of_measure[0]]) == "EA":
                    print (input_file, ":", sheet_name, "row", df.loc[row_index,row_num[0]], "-- Check Unit of Measure for wire")

#########################################################################################################################
### Normalize all of the column headers from all the BOM files so that they are the same.
#########################################################################################################################

def normalize_column_headers(header):
    
#### Here are those pesky global variables again.  

    global row_num, xp_pn, mfg_pn, quantity, description, level_val, ref_des
    global unit_of_measure, manufacturer, comp_text, rev_level
    global xp_sort_key_col_name, mfg_sort_key_col_name, file_col_name, final_column_order
    global row_num_found, xp_pn_found,  mfg_pn_found, quantity_found,description_found, level_found, ref_des_found
    global unit_of_measure_found,  manufacturer_found, comp_text_found, rev_level_found
    global cost_bom_tab_names
    
### These are the names of the columns to parse for in the various BOMs:
 
    header = str(header)
    header = header.lower()
#    print("header is", header)
    
    if header in row_num and not row_num_found:
        row_num_found = True   
#        print("Found row number field")
        return(row_num[0])
        
    if header in xp_pn and not xp_pn_found: 
        xp_pn_found = True
#        print ("Found XP PN field")
        return(xp_pn[0])
    
    if header in mfg_pn and not mfg_pn_found:
        mfg_pn_found = True
#        print ("found Mfg PN field")
        return(mfg_pn[0])
    
    if header in quantity and not quantity_found:
        quantity_found = True
#       print ("found quantity field")
        return(quantity[0])

    if header in description and not description_found:
        description_found = True
#        print("found description field")
        return(description[0])
        
    if header in level_val and not level_found: 
        level_found = True    
#        print ("found BOM level field")
        return(level_val[0])
    
    if header in unit_of_measure and not unit_of_measure_found:
        unit_of_measure_found = True
#        print ("found unit of measure field")
        return(unit_of_measure[0])
    
    if header in ref_des and not ref_des_found:
        ref_des_found = True
#        print ("found RefDes field")
        return(ref_des[0])
    
    if header in comp_text and not comp_text_found:
        comp_text_found = True
#        print ("found como_text field")
        return(comp_text[0])
    
    if header in rev_level and not rev_level_found:
        rev_level_found = True
#        print ("found rev_level field")
        return(rev_level[0])    
    

### This is the order of the columns in the unified BOMs.  
### Just realized this is dumb place for this...  

    final_column_order = [file_col_name, xp_sort_key_col_name, mfg_sort_key_col_name, row_num[0], level_val[0], xp_pn[0], mfg_pn[0], rev_level[0], quantity[0], description[0], comp_text[0],ref_des[0], unit_of_measure[0]]

#########################################################################################################################    
### This is the main BOM normalization function.  It should really be broken down into more subroutines
#########################################################################################################################

def bom_norm(input_file, output_file_directory):
    

### Here are those pesky global variables again.  :)

    global row_num_found, xp_pn_found,  mfg_pn_found, quantity_found,description_found, level_found, ref_des_found
    global unit_of_measure_found,  manufacturer_found, comp_text_found, rev_level_found
    global cost_bom_tab_names
    
    global row_num, xp_pn, mfg_pn, quantity, description, level_val, ref_des
    global unit_of_measure, manufacturer, comp_text, rev_level
    global xp_sort_key_col_name, mfg_sort_key_col_name, file_col_name, final_column_order
    

### I will make this a stand-alone routine once I get the global variable situation figured out.
### This figures out the BOM type based on the file name.  
### Eventually, I want to have in the configuration file the possible strings in the file names and 
### how it maps to BOM type and what string to look for in the top left cell of the BOM data.
### Hard code for now.


#    print("In BOM Norm: input file name = ",input_file)
    if "solidworks" in input_file.lower():
        bom_type = "solidworks"
        top_left_bom_cell = "LEVEL"
    elif "altium" in input_file.lower():
        bom_type = "altium"
        top_left_bom_cell = "No."
    elif "4th" in input_file.lower():
        bom_type = "4thshift"
        top_left_bom_cell = "EffectivityDate"
    elif "001-" in input_file.lower():
        bom_type = "WIP_Bom"
        top_left_bom_cell = "No."
    elif "s4" in input_file.lower():
        bom_type = "S4"
        top_left_bom_cell = "Explosion level"
## ADDED VNM TO FOLLOWING LINE on 8/10/20:
    elif "vn" in input_file.lower() or "vtn" in input_file.lower() or "vtm" in input_file.lower() or "vnm" in input_file.lower():
        bom_type = "VN"
        top_left_bom_cell = "No."
    else:

### I haven't come up with an exception handling strategy yet... :)

        bom_type = "Not Recognized"
     
#    print("BOM type is", bom_type, "   cell to look for is", top_left_bom_cell)


#########################################################################################################################    
### This should be routine the beginning of a routine that iterates through the sheets of the file and then calls a routine to parse it
#########################################################################################################################

    first_tab_of_input_bom = True
    delete_range = 0
    
## Create output file name
    
    path = Path (input_file)
    filename_wo_ext = path.stem
#    print("file name without extension is", filename_wo_ext)
    output_file_name = output_file_directory+"/"+filename_wo_ext+"_norm.csv"
    
    input_file_name = input_file

## Read each sheet of the input file  
    xls = pd.ExcelFile(input_file)
    sheets = xls.sheet_names
    for sheet_name in sheets:

#########################################################################################################################    
### This should be a routine to process the sheets
#########################################################################################################################
  
#
#        print("Sheet name is:", sheet_name)
        df = pd.read_excel(xls, sheet_name, header=None) 

## The BOM types VN and WIP_Bom have a tab that has the BOM data in it.
## Ideally the mapping between the BOM types and the possible BOM tab names would be in the configuration file.
## For now, hard code it.
        
        if  ((bom_type =="VN") or (bom_type == "WIP_Bom")) and (sheet_name not in cost_bom_tab_names):
            really_bad_programming_practice = 42
#            print("Skip ", input_file, ":", sheet_name)
        else:
            
## Change all "nan" fields to ""
            df = df.fillna("")
            
## Initiatliaze variables for each sheet   
## Here are those global variables again....
            
            parent_level_string = ""
            level_adjustment = 0
            
            row_num_found = False
            xp_pn_found = False
            mfg_pn_found = False
            quantity_found = False
            description_found = False
            level_found = False
            xp_pn_found = False
            ref_des_found = False
            unit_of_measure_found = False
            manufacturer_found = False
            comp_text_found = False
            rev_level_found = False
    
## Read the top rows to figure out what BOM level the data should correspond to and where the data starts
            
#            print("df is", df)
#            print("len(df)) = ", len(df))
#            print("range(len(df))) = ", range(len(df)))
            
            delete_range = -999

            for row in range(len(df)):
#                print ("row is", row)
                value_to_check = df.iloc[row,0]
 #               print("value to check = ", value_to_check)
                
                if(value_to_check != ""):
                    
## If there is a "Parent Levels" row, use that string to pre-pend the part numbers in the sort key.
## Used for Solidworks BOMs, but could be used for any BOM.

                    if (value_to_check =="Parent Levels" ):
                        parent_level_string = str(df.iloc[row,1])
    #                    print("parent level string =" ,parent_level_string )
                        parent_bom_level = parent_level_string.count("|")+1
    #                    print ("bom_level_count = ", parent_bom_level)
                        
## If there is a "Level Adjustment" row, use this as an offset to the level values in this sheet.
## Used for Solidworks BOMs, but could be used for any BOM.
                        
                    elif(value_to_check == "Level Adjustment"):
                        level_adjustment = df.iloc[row,1]

## Look for the top left cell value depending on the BOM type.  (See around line 252 of this file :).  
## This is used to figure out how many rows to delete at the top of the BOM to get to the BOM fields.
## (The WIP and VN BOMs have a lot of extraneous rows at the top of the actual BOM data.)
                    elif (value_to_check == top_left_bom_cell ):
#                        print("found row header")
                        delete_range = row+1
#                        print("delete range value = ", delete_range)
                        break

### If the top left cell is not found, then the BOM is not formatted correctly to be processed.                    
#            print("delete range value before if = ", delete_range)
            if delete_range == -999:
                print(input_file_name,":",sheet_name,"is not formatted for BOM Compare.  Delete extraneous and hidden tabs.  Fix file and re-run")
                print("Cannot find top left bom cell value", top_left_bom_cell)
                
###  Not sure what "quit()" does. Just tried it as a way to bail out :)
                
                quit()
           
    #        print("delete range = 0 to ",delete_range)
    #        print("Parent levels = ", parent_level_string)

                
## Re-assign the column header based on the top row
            df.columns = df.iloc[delete_range-1]
## I think this is not needed.  I will delete this when I have time to make sure it works.  :)
            column_headers = df.columns
    
## Delete the rows above the column headers
            df = df.drop(df.index[0:delete_range])
            column_headers = df.columns
#            print("column header is", df.columns)
            
## Make the column headers strings (Needed for parsing Cost BOM)
    
            df.iloc[1:].to_string(header=False, index=False)
    
    #        print ("iloc", df.iloc[1,:])
    
## Look at each column header and normalize it
            count = 0
            for i in column_headers:
                new_value = normalize_column_headers (df.columns[count])
                df.columns.values[count] = new_value
                count = count + 1
            
        #    column_headers = df.columns
    
## Add a "sort key" column and fill in the values
            df.insert(2,xp_sort_key_col_name, "")
            df.insert(2,mfg_sort_key_col_name, "")
            
## Add a "File" column as the first column.  
            df.insert(0,'File', "")
            
## Add any missing columns
    
            if not row_num_found:
        #        print("add row num col")
                row_num_found = True
                df.insert(2, row_num[0],"") 
    
            if not xp_pn_found:
        #        print("add xp pn col")
                xp_pn_found = True
                df.insert(2, xp_pn[0],"") 
    
            if not ref_des_found:
        #        print("add ref_des col")
                ref_des_found = True
                df.insert(2, ref_des[0],"")   
                
            if not comp_text_found:
        #        print("add text col")
                comp_text_found = True
                df.insert(2, comp_text[0],"")      
              
            if not rev_level_found:
        #        print("add rev_level col")
                rev_level_found = True
                df.insert(2, rev_level[0],"") 
                
            if not level_found:
        #       print("add BOM level")
                level_found = True
                df.insert(2, level_val[0],"") 
                
            if not xp_pn_found:
        #       print("XP PN")
                xp_pn_found = True
                df.insert(2, xp_pn[0],"") 
                
            if not mfg_pn_found:
        #       print("Add Mfg PN")
                mfg_pn_found = True
                df.insert(2, mfg_pn[0],"") 
                
            if not unit_of_measure_found:
        #       print("Add Unit of Measure")
                unit_of_measure_found = True
                df.insert(2, unit_of_measure[0],"") 
                
            if not quantity_found:
                print("quantity not found.  BOM cannot be normalized")
                
            if not description_found:
                print("description not found.  BOM cannot be normalized")

#########################################################################################################################    
### This should be a separate function to add the sort keys
#########################################################################################################################
    
    
## Iterate through each row to fill in the sort key values
## Note: current_levels not used for Solidworks BOMs
            
            current_levels = ["","","","","", "", ""]
            
            first_solidworks_row_on_sheet = True 
            
            for row_index, row in df.iterrows():
##############################################################################                
##  Add sort key for Altium BOMs.  The Sort Key will be based on the XP P/N
##############################################################################
                
    #            print("Before BOM-specific logic")
                df.loc[row_index,file_col_name] = input_file+":"+sheet_name
                if bom_type == "altium":
                    part_num = str(df.loc[row_index,xp_pn[0]])
    #               print("row index = ",row_index, "xp part num =", xp_part_num)
    #                print("in altium part of loop")
                    df.loc[row_index,level_val[0]] = parent_bom_level + 1
                    if(part_num != ""):
                        if (parent_bom_level == 1):
                            sort_key = parent_level_string+"|"+str(part_num)+"||"
                        elif(parent_bom_level == 2):
                            sort_key = parent_level_string+"|"+str(part_num)+"|"
                        elif(parent_bom_level == 3):
                            sort_key = parent_level_string+"|"+str(part_num)
                        else:
                            print("ERROR - TOOL DOES NOT SUPPORT ",parent_bom_level, " BOMS")
                            break
        #                print ("Sort key is", sort_key)
                        df.loc[row_index,xp_sort_key_col_name] = sort_key
                        
                        
##############################################################################                
##  Add sort key for Solidworks BOMs  The Sort Key will be based on the XP P/N
##############################################################################           
                elif bom_type == "solidworks":

#                    print("in solidworks part of loop")
                    part_num = str(df.loc[row_index,xp_pn[0]])
#                    print("row index = ",row_index, "part num =", xp_part_num)
                    #if(part_num != ""):
                    if not pd.isna([part_num]):
#                        print ("BOM Level is ", df.loc[row_index,level_val[0]])
                        if  df.loc [row_index,level_val[0]] != "":
                            df.loc [row_index,level_val[0]] = df.loc[row_index,level_val[0]] + level_adjustment                                 
                            current_bom_level = df.loc[row_index,level_val[0]]
                        else:
                            print("MISSING BOM LEVEL - FIX AND RE-RUN")
                            break
                        
    #                   print ("BOM Level is ", df.loc[row_index,level_val[0]])
    
                        current_levels[current_bom_level-1] = xp_part_num
                        for i in range(current_bom_level,4):
                            current_levels[i]= ""
                        sort_key = current_levels[0] + '|' + current_levels[1] + '|' + current_levels[2] + '|' + current_levels[3]
    #                    print("sort key is", sort_key)
                        df.loc[row_index,xp_sort_key_col_name] = sort_key
                        
##############################################################################                
##  Add sort key for S4 BOMs  The Sort Key will be based on the Mfg P/N
##  It is largely the same as for the 4thShift BOM.  Make this more efficient.
##############################################################################     
                                  
                elif bom_type == "S4":
    
                    part_num = str(df.loc[row_index,mfg_pn[0]])
    #                print ("part num is", part_num)
    #                print("data frame is")
    #                print(df.to_string(index=False))
    #                print("end")
                    bom_level_string = str(df.loc[row_index, level_val[0]])
    #                print ("bom level string=", bom_level_string)
                    bom_level_string = bom_level_string[-1]
    #                print ("new bom_level = ", bom_level_string)
                    df.loc[row_index,level_val[0]] = int(bom_level_string[-1])
    
#  This was in the code twice... Test and remove                    part_num = df.loc[row_index,mfg_pn[0]]
    #                print ("part_num = ", part_num)
                    
    #               print("row index = ",row_index,  "part num =", part_num)
                    if part_num != "":
                        
                        current_bom_level = df.loc[row_index,level_val[0]]
    #                   print ("BOM Level is ", df.loc[row_index,level_val[0]])
    
                        current_levels[current_bom_level-1] = part_num
                        for i in range(current_bom_level,4):
                            current_levels[i]= ""
                        sort_key = current_levels[0] + '|' + current_levels[1] + '|' + current_levels[2] + '|' + current_levels[3]
    #                    print("sort key is", sort_key)
                        df.loc[row_index,mfg_sort_key_col_name] = sort_key 
 
##############################################################################                
##  Add sort key for 4thShift BOMs.  Sort Key will be based on XP P/N
##############################################################################     
                 
                elif bom_type == "4thshift":
                    part_num = str(df.loc[row_index,xp_pn[0]])
    #               print("row index = ",row_index, "part num =", part_num)
                    bom_level_string = str(df.loc[row_index,level_val[0]])
                    df.loc[row_index,level_val[0]] = int(bom_level_string[-1])
    
                    if(part_num != ""):   
    
                        current_bom_level = df.loc[row_index,level_val[0]]
    #                  print ("BOM Level is ", df.loc[row_index,level_val[0]])
    
                        current_levels[current_bom_level-1] = part_num
                        for i in range(current_bom_level,4):
                            current_levels[i]= ""
                        sort_key = current_levels[0] + '|' + current_levels[1] + '|' + current_levels[2] + '|' + current_levels[3]
    #                    print("sort key is", sort_key)
                        df.loc[row_index,xp_sort_key_col_name] = sort_key
                    
##############################################################################                
##  Add sort key for WIP and VN BOMs - for both XP P/N and Mfg P/N
##  Make this more efficient.
##############################################################################     
                          
                elif bom_type == "WIP_Bom" or bom_type == "VN":
#### THIS IS WHERE THE PART NUMBER IS SET FOR THE EXCEL BOM - EITHER MFG PN OR THE XP PART NUMBER
                    mfg_part_num = str(df.loc[row_index,mfg_pn[0]])
#                    print("row index = ",row_index, "mfg part num =", mfg_part_num)
                    if(mfg_part_num != ""): 
#                        print("part_num is not null")
                        if df.loc[row_index,level_val[0]] != "": 
#                            print("Level is not null")
                            current_bom_level = int(df.loc[row_index,level_val[0]])
#                            print ("BOM Level is ", df.loc[row_index,level_val[0]])
                            current_levels[current_bom_level-1] = mfg_part_num
                            for i in range(current_bom_level,4):
                                current_levels[i]= ""
                            sort_key = current_levels[0] + '|' + current_levels[1] + '|' + current_levels[2] + '|' + current_levels[3]
        #                    print("mfg sort key is", sort_key)
                            df.loc[row_index,mfg_sort_key_col_name] = sort_key
                        else:
                            print("BOM level is missing.  Fix file and re-run")
                            break
                        
                    xp_part_num = str(df.loc[row_index,xp_pn[0]])
#                    print("row index = ",row_index, "XP part num =", xp_part_num)
                    if(xp_part_num != ""): 
#                        print("xp part_num is not null")
                        if df.loc[row_index,level_val[0]] != "": 
#                            print("Level is not null")
                            current_bom_level = int(df.loc[row_index,level_val[0]])
#                            print ("BOM Level is ", df.loc[row_index,level_val[0]])
                            current_levels[current_bom_level-1] = xp_part_num
                            for i in range(current_bom_level,4):
                                current_levels[i]= ""
                            sort_key = current_levels[0] + '|' + current_levels[1] + '|' + current_levels[2] + '|' + current_levels[3]
        #                    print("sort key is", sort_key)
                            df.loc[row_index,xp_sort_key_col_name] = sort_key
                        else:
                            print("BOM level is missing.  Fix file and re-run")
                            break

                else:
                    
##############################################################################                
##  Add this section just in case it is needed.  
##  How do I bail out from here?  :)
##############################################################################     
                    print ("CODE NOT YET IMPLEMENTED FOR THIS BOM TYPE")
    

        
    #        print("columns are ", df.columns)         
            column_headers = df.columns
    #       print("column headers are", column_headers)
            
##############################################################################                
##  Figure out if all required column headers have been found.
##  Ideally the configuration file should note which columns are required.
##############################################################################

            all_required_col_headers = (xp_pn_found or mfg_pn_found) and quantity_found and description_found and level_found
        
            if not all_required_col_headers: 
                print("NOT ALL required column headers present")
    #            print (column_headers)
                print("file", input_file,"is not normalized)")

##############################################################################                
##  I was thinking that the routine should return True or False depending on if the BOM was normalized.
##  But I don't know if it actually works, and I didn't take time to figure it out.  :)
##############################################################################
                return True
    #        else:
 
#Put the columns in order
                print("all column headers found")
                
    #        print("df before re-order", df)
            df = df[final_column_order]

### Ya.  This is called in the wrong place.  I will move it once I tidy up the rest of the code.
            check_bom_fields(input_file, sheet_name, bom_type, df, final_column_order)
            
##############################################################################                
##  Combine all of the data from each sheet of the file in to a csv file and write to output directory.
##############################################################################

            if first_tab_of_input_bom:

                df.to_csv(output_file_name, index=False, header=True)
                first_tab_of_input_bom = False
            else:

                df.to_csv(output_file_name, index = False, mode='a', header=False)
    
            print("Normalized>>", input_file)
            return_error = False

###########################################################################
### Main Program Starts Here
###########################################################################
def normalize_boms (orig_bom_directory, output_file_directory):
    
    

#    print ("orig_bom_directory is", orig_bom_directory, " output_file_diretory = ", output_file_directory)
 
    initialize_data()

    
    for filename in os.listdir(orig_bom_directory):
        
        file_with_path = orig_bom_directory + "/" + filename
#        print ("file with path = ", file_with_path)
#        print ("output file = ", output_file_directory )
        if filename.endswith(".xls") or filename.endswith(".xlsx"):
## ADDED NEXT LINE ON 8/10/20 
            print("")
            print("Processing file:  ", filename)
            norm_error = bom_norm(file_with_path, output_file_directory)
            if norm_error:
                print("Error - file cannot be normalized")
                break

    print("")
    print("NEXT STEP:  If no errors were reported, run the 'Unify Norm Boms' program on the output directory")
# The following makes this program start running at normalize_all() 
# when executed as a stand-alone program.    
if __name__ == '__main__':
    if len(sys.argv) != 3:
        print ("ERROR:  Re-run program with input_bom and output_bom directories")
        sys.exit(1)
    else:
        normalize_boms(sys.argv[1], sys.argv[2])

   
    
