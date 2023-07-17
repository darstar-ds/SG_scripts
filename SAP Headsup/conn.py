import pandas as pd
import pyodbc

# Establish a connection to the database
conn = pyodbc.connect("DRIVER={SQL Server};SERVER=boczek;DATABASE=sgdb;Trusted_Connection=yes;")

# Define your SQL query
query = "select product as Project, id as SubProject, falcon_phase as Step, return_date as ReturnDate, request_description as Volume, SL.short_code as Source, TL.language_tag as Target, jp_environment AS Environment from JobParts INNER JOIN Languages AS SL (nolock) ON (JobParts.jp_src_lang_id = SL.language_id) INNER JOIN Languages AS TL (nolock) ON (JobParts.jp_trg_lang_id = TL.language_id) where platform='SAP' and return_date > DATEADD(day, 7, GETDATE()) order by job_part_id desc"

# Execute the SQL query and store the results in a DataFrame
df = pd.read_sql_query(query, conn)

# Close the database connection
conn.close()

# Print the DataFrame
print(df[["Volume", "Target", "SubProject"]])
