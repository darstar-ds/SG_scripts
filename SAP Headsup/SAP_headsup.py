import pypyodbc 
import pandas as pd

cnxn = pypyodbc.connect(
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=boczek;"
    "Database=sgdb;"
    # "uid=login;pwd=pass"
                        )
df = pd.read_sql_query(
    """
    SELECT 
    product as Project 
    , id as SubProject
    , falcon_phase as Step
    , return_date as ReturnDate
    , request_description as Volume
    , SL.short_code as Source
    , TL.language_tag as Target
    , jp_environment as Environment 
    from JobParts 
    INNER JOIN Languages AS SL (nolock) ON (JobParts.jp_src_lang_id = SL.language_id) 
    INNER JOIN Languages AS TL (nolock) ON (JobParts.jp_trg_lang_id = TL.language_id) 
    where platform="SAP" and return_date > DATEADD(day, 7, GETDATE()) 
    order by job_part_id desc
    """, cnxn)
