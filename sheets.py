#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Nov 13 10:38:22 2017

@author: Home

Google Sheets Project

Version 1
To be included:
1. Read from a specified google worksheet within a sheet
2. write a full table to a worksheet within a sheet


"""

#file path is to the location of the json file from google
def auth(file_path):
    from oauth2client.service_account import ServiceAccountCredentials
    scope = ['https://spreadsheets.google.com/feeds']
    creds = ServiceAccountCredentials.from_json_keyfile_name(file_path, scope)
    return creds


def read(sheet_name,worksheet_name,creds):
    import gspread
    import pandas as pd

    #Authorizations and sheet selection
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name)
    worksheet = sheet.worksheet(worksheet_name)

    #gets all data from the worksheet in the form of a dict
    data = worksheet.get_all_records()
    data = pd.DataFrame.from_dict(data)
    
    #Pandas automatically re-orders columns alphabetically
    #In-order to bring in data in the same format one needs to bring in the row by itself and use it to arrange the df
    rowcol_list = worksheet.row_values(1)

    col_list = []
    for i in range(0, len(data.columns)):
        col_list.append(rowcol_list[i])
    
    return data[col_list] 


def write(data, sheet_name,worksheet_name,creds):
    import gspread
    import pandas as pd
    
    #Authorizations and sheet selection
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name)
    worksheet = sheet.worksheet(worksheet_name)
     
    #Gspread doesn't account for column names
    #This creates a new df to add to the top of the original df
    cols = data.columns 
    col_names = {}
    for i in cols:
        col_names[i] = i        
    insert_names = pd.DataFrame(col_names, index = [0])
    data = pd.concat([insert_names, data]).reset_index(drop = True)
    data = data[cols]
    
    #creates a list of all the necessary cells based on the size of the df    
    cell_range = '{col_i}{row_i}:{col_f}{row_f}'.format(
            col_i=chr((0) + ord('A')),    # converts number to letter
            col_f=chr((len(cols)-1) + ord('A')),      # subtract 1 because of 0-indexing
            row_i=1,
            row_f=len(data))
    cell_list = worksheet.range(cell_range)

    #Transforms data into one list to work with cell_range
    values = []
    for q in range(0, len(data)):
        for i in range(0, len(cols)):
            values.append(data[cols[i]][q])
        
    for i, val in enumerate(values):  
        cell_list[i].value = val         

    #Worksheets need to be able to accomodate all cells that are part of an update or it will fail.
    #This re-szes the worksheet ahead of the update.
    worksheet.resize(len(data), len(data.columns))
    

    #I don't think gspread can't send more than 50,000 cells at once during a multi cell sheet update.
    #This creates bins of 45,000 cells to send off, and when a == len(cell_list) the remaining bin that didn't reach 45,000 cells will be sent
    chunk = []
    a = 0
    for i in range(0, len(cell_list)):
        if len(chunk) < 45000:
            chunk.append(cell_list[i])
            a += 1
            if a == len(cell_list):
                sheet.update_cells(chunk)            
        elif len(chunk) == 45000:
            sheet.update_cells(chunk)
            chunk = []
            a += 1
            
    return print('Upload to Worksheet:%r within Sheet:%r Complete'  %(worksheet_name, sheet_name))

        



