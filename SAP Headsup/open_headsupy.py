import os
import time
import xlwings as xw 

CURR_PATH = "d:\\Moje dokumenty\\SG_scripts_data\\SAP\\headsupy\\"
JOBPARTS_EXPORT = CURR_PATH + "SAPowe_Headsupy.xlsx"
UPDATED_EXPORT = CURR_PATH + "SAPowe_Headsupy_updated.xlsx"
SAVE_PATH_UP = "r" + "'" + UPDATED_EXPORT + "'"

# function to close a workbook given name
def close_wb(wbname):
    
    # try: 
        app = xw.apps.active # get the active Excel application
        wb = app.books[wbname] # make workbook with given name active
        print('saving workbook',wbname)
        wb.save()
        time.sleep(5)
        print('closing workbook',wbname)
        wb.close()
    # except: pass


# os.system('start "excel" "' + JOBPARTS_EXPORT + '"')

os.startfile(JOBPARTS_EXPORT)
time.sleep(10)
close_wb("SAPowe_Headsupy.xlsx")