import pandas as pd
import os

CURR_PATH = "./Reuter/2023-05-23/"
EXPORTS4TRANS = CURR_PATH + "exports4trans/"
EXPORTED_COLUMNS = CURR_PATH + "4trans_sheets/"
TRANS_COLUMNS = [
        'Kurzbeschreibung', 
        'Optionstext',
        'Lieferumfang', 
        # 'Einleitung',
        'Langbeschreibung', 
        'Ergaenzung-Bullets', 
        'Weitere-Anmerkungen', 
        'Weitere-Besonderheiten',
        'Variante',
        'Hinweis',
        ]

def find_unique_4trans(path, file_name):
    xls = pd.ExcelFile(path + file_name)
    # print(xls.sheet_names)  # see all sheet names
    xls_tabs = xls.sheet_names
    
    for tab_name in xls_tabs:
        print(f"Zakladka: {tab_name}")
        df_1sheet = xls.parse(tab_name)  # read a specific sheet to DataFrame
        print(df_1sheet.head())
        print(df_1sheet.shape)
        for column in TRANS_COLUMNS:
            print(f"Kolumna: {column}")
            df_1sheet_deu = df_1sheet[df_1sheet["Sprache"]=="deu"]
            print(df_1sheet_deu.head())
            df_1column = df_1sheet_deu.loc[:, column]
            print(f"Liczba unikalnych wystapień w kolumnie {column}: {df_1column.nunique()}")
            nparray_1column_unique = df_1column.unique()
            print(f"Typ danych: {type(nparray_1column_unique)}")
            spreadsheet_name = EXPORTED_COLUMNS + file_name + "_" + tab_name + "_" + column + ".xlsx"
            df_1column_unique = pd.DataFrame(nparray_1column_unique) # convert numpy array into dataframe
            df_1column_unique.columns = ["DE"]
            df_1column_unique["PL"] = None
            print(df_1column_unique.head())
            df_1column_unique.to_excel(spreadsheet_name, sheet_name=column, index=False)
            

dir_list = os.listdir(EXPORTS4TRANS)

for file in dir_list:
    find_unique_4trans(EXPORTS4TRANS, file)
