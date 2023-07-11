import pandas as pd
import os
import openpyxl

CURR_PATH = "d:\\Moje dokumenty\\SG_scripts_data\\SAP\\headsupy\\"
JOBPARTS_EXPORT = CURR_PATH + "SAPowe_Headsupy.xlsx"


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
    df['Volume_units'] = new4[1]


    df.drop("Volume", inplace=True, axis=1)

    return df


def repair_nulls(df):

    for nan_index, nan_row in df.iterrows():
        # print(type(nan_row["Volume"]))
        print(f"Nan_index: {nan_index}; Volume type: {type(nan_row['Volume'])}")
        if type(nan_row["Volume"]) == float:
            NaN_Project = nan_row["Project"]
            NaN_SubProject = nan_row["SubProject"]
            NaN_Step = nan_row["Step"]
            print(f"Nan_Project: {NaN_Project}; Nan_SubProject: {NaN_SubProject}; Nan_Step: {NaN_Step}")
            sim_rows = df.loc[
                (df["Project"] == NaN_Project)\
                & (df["SubProject"] == NaN_SubProject)\
                & (df["Step"] == NaN_Step)\
                # looking for a row similar to the one with Nan values
                ]
            # print(type(sim_rows))
            print(f"Sim_rows shape: {sim_rows.shape}")
            print(sim_rows[["Volume", "Target"]]) # "Project", "SubProject", "Step",

            for full_index, full_row in sim_rows.iterrows():
                if type(full_row["Volume"]) == str:
                    print(f"Full_index: {full_index}; Full non NaN row: \n {full_row[['Volume', 'Target']]}")
                    V_task = full_row["Volume"]
                    # V_no = full_row["Volume_no"]
                    # V_units = full_row["Volume_units"]
                    print(f"V_task: {V_task}")
                    df.at[nan_index, "Volume"] = V_task
                    # nan_row["Volume_no"] = V_no
                    # nan_row["Volume_units"] = V_units
                    print(f"Nan_index: {nan_index}; Full repaired row \n {df.iloc[nan_index]}")
                else:
                    pass
                       
            print(f"Nan_index: {nan_index}; Volume: {nan_row['Volume']}")
            # print(list(df.index.values))
            # print(len(df.index))
        else:
            pass
        # print(df)   
    return df

jobparts_import = pd.ExcelFile(JOBPARTS_EXPORT)
df_jobparts_import = jobparts_import.parse() # convert into dataframe
print(df_jobparts_import.columns)
print(df_jobparts_import.shape)
print(df_jobparts_import['Step'].unique())

# select rows with DTP tasks only and reset index
df_DTP_headsups = df_jobparts_import.loc[df_jobparts_import['Step'].isin(['FOR_CREPDF', 'DTPCREAT', 'IMPLMT'])]
df_DTP_headsups.reset_index(inplace=True, drop=True)

# replace null values wiith data from similar tasks
df_DTP_headsups_repaired = repair_nulls(df_DTP_headsups)

# split Volume column into task-number-units
df_Volume_split = split_Volume(df_DTP_headsups_repaired)

# remove rows with no task data
df_Volume_split.dropna(inplace=True)
print(df_Volume_split.columns)
print(df_Volume_split.info())