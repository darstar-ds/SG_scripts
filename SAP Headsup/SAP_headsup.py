import pandas as pd
import os
import time
import tkinter as tk
import xlwings as xw 
from datetime import datetime, timedelta
import streamlit as st
import pyodbc

CURR_PATH = "d:\\Moje dokumenty\\SG_scripts_data\\SAP\\headsupy\\"
JOBPARTS_XLS = "SAPowe_Headsupy.xlsx"
JOBPARTS_EXPORT = CURR_PATH + JOBPARTS_XLS


def import_data():
    # Establish a connection to the database
    conn = pyodbc.connect("DRIVER={SQL Server};SERVER=boczek;DATABASE=sgdb;Trusted_Connection=yes;")

    # Define your SQL query
    query = "select product as Project, id as SubProject, falcon_phase as Step, return_date as ReturnDate, request_description as Volume, SL.short_code as Source, TL.language_tag as Target, jp_environment AS Environment from JobParts INNER JOIN Languages AS SL (nolock) ON (JobParts.jp_src_lang_id = SL.language_id) INNER JOIN Languages AS TL (nolock) ON (JobParts.jp_trg_lang_id = TL.language_id) where platform='SAP' and return_date > DATEADD(day, 7, GETDATE()) order by job_part_id desc"

    # Execute the SQL query and store the results in a DataFrame
    df = pd.read_sql_query(query, conn)

    # Close the database connection
    conn.close()
    return df

def show_timer(seconds):
    root = tk.Tk()
    label = tk.Label(root, font=("Arial", 24), width=10)
    label.pack()

    for i in range(seconds, 0, -1):
        label["text"] = f"Timer: {i}s"
        root.update()
        time.sleep(1)

    root.destroy()

def update_headsupy(wb_path):
    wbname = os.path.basename(wb_path)
    os.startfile(wb_path)

    show_timer(10)
    # time.sleep(10)
    
    app = xw.apps.active # get the active Excel application
    wb = app.books[wbname] # make workbook with given name active
    print('saving workbook',wbname)
    wb.save()
    time.sleep(3)
    print('closing workbook',wbname)
    wb.close()

def split_Volume(df):
    '''
    Function takes dataframe and
    splits "Volume" column into three values:
    volume_taskname
    volume_no
    volume_units
    Returns dataframe with three new columns and Volume column removed
    '''
    volume_column = df['Volume']

    new1 = volume_column.str.split("\n", n = 1, expand = True) # remove rubish from the beginning
    rubish1 = new1[0]
    new_volume1 = new1[1]
    new2 = new_volume1.str.split(";", n = 1, expand = True) # remove rubish from the end
    rubish2 = new2[1]
    new_volume2 = new2[0]
    new3 = new_volume2.str.split(": ", n = 1, expand = True) # divide task name and scope
    df['Volume_task'] = new3[0]
    scope = new3[1]
    new4 = scope.str.split(" ", n = 1, expand = True) # divide number and units
    df['Volume_no'] = new4[0]
    df['Volume_units'] = new4[1].astype(str)

    df.drop("Volume", inplace=True, axis=1)

    return df


def repair_nulls(df):
    '''
    The function takes a dataframe
    finds a row with NaN value (float type)
    selects similar rows with filledin values
    inserts the value from similar row
    returns dataframe
    '''
    for nan_index, nan_row in df.iterrows():
        
        # print(f"Nan_index: {nan_index}; Volume type: {type(nan_row['Volume'])}")
        if type(nan_row["Volume"]) == float:
            NaN_Project = nan_row["Project"]
            NaN_SubProject = nan_row["SubProject"]
            NaN_Step = nan_row["Step"]
            # print(f"Nan_Project: {NaN_Project}; Nan_SubProject: {NaN_SubProject}; Nan_Step: {NaN_Step}")
            
            # looking for a row similar to the one with Nan values
            sim_rows = df.loc[
                (df["Project"] == NaN_Project)\
                & (df["SubProject"] == NaN_SubProject)\
                & (df["Step"] == NaN_Step)\
                ]
            
            # looking for a filledin values among similar rows
            for _, full_row in sim_rows.iterrows():
                if type(full_row["Volume"]) == str:
                    # print(f"Full_index: {full_index}; Full non NaN row: \n {full_row[['Volume', 'Target']]}")
                    V_task = full_row["Volume"]
                    # print(f"V_task: {V_task}")
                    df.at[nan_index, "Volume"] = V_task
                    # print(f"Nan_index: {nan_index}; Full repaired row \n {df.iloc[nan_index]}")
                else:
                    pass
                       
            # print(f"Nan_index: {nan_index}; Volume: {nan_row['Volume']}")
            
        else:
            pass   
    return df

def add_prevdays(df):
    df["ReturnDate"] = (pd.to_datetime(df["ReturnDate"]).dt.date).astype('datetime64')
    df["ReturnDate-1"] = df["ReturnDate"] - timedelta(days = 1)
    df["ReturnDate-2"] = df["ReturnDate"] - timedelta(days = 2)
    df["ReturnDate-3"] = df["ReturnDate"] - timedelta(days = 3)
    df["ReturnDate-4"] = df["ReturnDate"] - timedelta(days = 4)
    # print(df.info())
    # print(df[["Volume_no", "ReturnDate", "ReturnDate-1", "ReturnDate-2", "ReturnDate-3", "ReturnDate-4"]].head)
    return df

def add_prevworkload(df):
    df["Volume_no"] = (df["Volume_no"]).astype(float)/5
    df["Volume_no-1"] = df["Volume_no"]
    df["Volume_no-2"] = df["Volume_no"]
    df["Volume_no-3"] = df["Volume_no"]
    df["Volume_no-4"] = df["Volume_no"]
    # print(df.info())
    # print(df[["Volume_no", "Volume_no-1", "Volume_no-2", "Volume_no-3", "Volume_no-4"]].head)
    return df

def into_dictionary(df):
    '''
    Takes dataframe and convert data into dict format
    {"date1": {"units": unit, "scope": sum}}
    '''
    scope_sum = {}
    for _, row in df.iterrows():
        r_date_vol_no = [
                        (row["ReturnDate-4"], row["Volume_no-4"]),
                        (row["ReturnDate-3"], row["Volume_no-3"]),
                        (row["ReturnDate-2"], row["Volume_no-2"]),
                        (row["ReturnDate-1"], row["Volume_no-1"]),
                        (row["ReturnDate"], row["Volume_no"])
                        ]
        
        for r_date, vol_no in r_date_vol_no:
            date = r_date
            scope = vol_no
            unit = row["Volume_units"]
            # print(f"Date: {date}, Scope: {scope} {unit}")
            
            if date not in scope_sum:
                # print(f"No key {date} in dict. Setting empty key.")
                scope_sum[date] = {'Frames': 0, 'Pages': 0, 'Hours': 0}
                scope_sum[date][unit] = scope
            elif date in scope_sum and unit in scope_sum[date]:
                scope_sum[date][unit] += scope
            else:
                scope_sum[date][unit] = scope
         
        # print(scope_sum)
    return scope_sum



# Changed my mind
# Code below imported Excel file
# Changed to direct import data from database

# update jobparts export with the newest data
# update_headsupy(JOBPARTS_EXPORT)

# import jobparts export to python dataframe
# jobparts_import = pd.ExcelFile(JOBPARTS_EXPORT)
# df_jobparts_import = jobparts_import.parse() # convert into dataframe


# import data from database
df_jobparts_import = import_data()
print(df_jobparts_import.columns)
print(df_jobparts_import.shape)
print(df_jobparts_import['Step'].unique())


# select rows with DTP tasks only and reset index
df_DTP_headsups = df_jobparts_import.loc[df_jobparts_import['Step'].isin(['FOR_CREPDF', 'DTPCREAT'])]
# , 'IMPLMT' - excluded as it doubles the number of pages/frames
df_DTP_headsups.reset_index(inplace=True, drop=True)

# replace null values wiith data from similar tasks
df_DTP_headsups_repaired = repair_nulls(df_DTP_headsups)

# split Volume column into task-number-units
df_Volume_split = split_Volume(df_DTP_headsups_repaired)

# remove rows with no task data
df_Volume_split.dropna(inplace=True)
# print(df_Volume_split.columns)
# print(df_Volume_split.info())

# add previous days columns
df_PrevDays_added = add_prevdays(df_Volume_split)

# add previous days workload
df_PrevDays_workload = add_prevworkload(df_PrevDays_added)

# convert dataframe to dictionary
dict_day_scope = into_dictionary(df_PrevDays_workload)

# covert dictionary to dataframe
dict_day_scope = {datetime.strptime(str(key), "%Y-%m-%d %H:%M:%S"): value for key, value in dict_day_scope.items()}
df_day_scope = pd.DataFrame.from_dict(dict_day_scope)
df_day_scope_trans = (df_day_scope.T).sort_index(ascending=True)
print(df_day_scope_trans.head())
df_day_scope_trans.to_csv(CURR_PATH + "workscope.csv")

# display data in streamlit
st.write("SAP marketing forecasts")
st.bar_chart(df_day_scope_trans)
st.dataframe(df_day_scope_trans)