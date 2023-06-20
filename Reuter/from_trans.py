import pandas as pd
import os
import openpyxl



def translated_gloss(trans_path):
    '''
    Provide the path to the directory
    with translated xlsx files.
    Empty dataframe "translated_gloss" is created and
    each file is added to the dataframe.
    '''
    dir_list = os.listdir(trans_path) 
    translated_gloss = pd.DataFrame(columns=["DE", "PL"])

    # get all translated files and build trans memmory as dataframe
    for file in dir_list:
        trans_file_path = trans_path + file
        one_trans_file = pd.ExcelFile(trans_file_path)
        df_one_trans_file = one_trans_file.parse() # convert into dataframe
        df_one_trans_file = df_one_trans_file[["DE", "PL"]] #select only DE and PL columns
        translated_gloss = pd.concat([translated_gloss, df_one_trans_file], axis=0) #add the current file to the gloss
        
    return translated_gloss

def translate_workbook(workbook_path, columns):
    '''
    Provide the path to xlsx file to be translated.
    Select the columns to be translated.
    '''
    # xls = pd.ExcelFile("./2023-01-03_Export_für_Hersteller_Übersetzung_PL.xlsx")
    wb = openpyxl.load_workbook(workbook_path)
    # xls_tabs = xls.sheetnames
    for tab in wb.sheetnames: #for each worksheet in a workbook
        ws = wb[tab]
        row_limit = ws.max_row
        print(f"Number of rows: {row_limit}")
        print(f"Current sheet: {ws}")
        print(f"Current tab: {tab}")
        # print(type(tab))
        for column in columns:
            curr_row = 2
            while curr_row < row_limit:
                deu_cell_1 = ws.cell(row=curr_row, column=column)
                deu_cell_2 = ws.cell(row=curr_row+1, column=column)
                pol_cell_1 = ws.cell(row=curr_row+2, column=column)
                pol_cell_2 = ws.cell(row=curr_row+3, column=column)
                
                if deu_cell_1.value is None:
                    pol_cell_1.value = None
                else:
                    pol_cell_1.value = transmem_dict[deu_cell_1.value]
                
                if deu_cell_2.value is None:
                    pol_cell_2.value = None
                else:
                    pol_cell_2.value = transmem_dict[deu_cell_2.value]
                
                curr_row += 4
        
        # remove tables, i.e. remove autofilter
        for item in ws.tables.items():
            del ws.tables[item[0]]
    return wb.save("./Reuter/2023-05-11_VB/2023-05-11_Export_DS_Riho.xlsx")





transmem = translated_gloss("./Reuter/2023-05-11_VB/translated/")
transmem.to_excel("./Reuter/slownik.xlsx", index=False)
transmem_dict = dict(transmem.values)
translated_xls = translate_workbook("./Reuter/2023-05-11_VB/2023-05-11_Export_Riho.xlsx", [6, 7, 9, 10, 11, 12, 13, 16, 17])