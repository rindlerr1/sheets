#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Nov 13 10:38:22 2017

@author: Home

Google Sheets Project

Version 1
To be included:
1. Read from a specified google sheet
2. write a full table to a sheet
    A. Needs to account for max number of cells which can be sent via API at once.
    B. Needs to account for automatic spreadsheet re-sizing


email: version1@sheet-v1.iam.gserviceaccount.com    
"""

email_message = ["You will need this email to be shared with your google sheet.",
          "email: version1@sheet-v1.iam.gserviceaccount.com"]
def email():
    for line in email_message:
        print(line)


def read(sheet_name):
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd

    scope = ['https://spreadsheets.google.com/feeds']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open(sheet_name).sheet1

    data = sheet.get_all_records()

    data = pd.DataFrame.from_dict(data)
    
    width = len(data.columns)

    import string
    alphabet = list(string.ascii_lowercase)

    cell_width = []
    for i in range(0, width):
        cell_width.append(alphabet[i]+str(1))

    big_list = sheet.row_values(1)

    col_list = []
    for i in range(0, len(cell_width)):
        col_list.append(big_list[i])
    
    return data[col_list] 


def write(data, sheet_name):
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd
    
    scope = ['https://spreadsheets.google.com/feeds']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).sheet1
        
    cols = data.columns 
    col_names = {}
    for i in cols:
        col_names[i] = i
        
    insert_names = pd.DataFrame(col_names, index = [0])
    data = pd.concat([insert_names, data]).reset_index(drop = True)
    data = data[cols]
    
    cell_range = '{col_i}{row_i}:{col_f}{row_f}'.format(
            col_i=chr((0) + ord('A')),    # converts number to letter
            col_f=chr((len(cols)-1) + ord('A')),      # subtract 1 because of 0-indexing
            row_i=1,
            row_f=len(data))

    cell_list = sheet.range(cell_range)

    values = []
    for q in range(0, len(data)):
        for i in range(0, len(cols)):
            values.append(data[cols[i]][q])
        
    for i, val in enumerate(values):  
        cell_list[i].value = val         


    sheet.resize(len(data), len(data.columns))
    
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
            
    return print('Upload to %r Complete'  %(sheet_name))

        



