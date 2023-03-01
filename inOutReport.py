import pandas as pd
import pyodbc
from datetime import time
import numpy as np
import os 

# ======================================= DB CONN & Retrieve =======================================================
server = '159.69.174.28,14333' 
database = 'AddDBName'
username = 'AddUserName'
password = 'AddPass' 
# sqlserver conn
try:
    conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password
    )
except Exception as e:
    print(e)

cursor = conn.cursor()

df_attlog = pd.read_sql(
    
						"""
							SELECT
								a.ID,
								a.DATE,
								a.TIME, 
								a.STATUS,
								a.DEVICE,
								a.NAME,
                                e.branchName,
								d.name as DEPARTMENT
							FROM
								ATTLOG a
							LEFT JOIN
								Employee e
							ON
								e.id=a.ID
							LEFT JOIN
								Departments d
							ON
								d.id=e.deparmentId

							WHERE
								a.DATE > '2022-11-30'
						
						""", conn
						
						)

# print(df_attlog)
# df_attlog.to_excel("dec_atten.xlsx")

#====================================== Data Cleaning ===============================================================
filter_unique_latest = df_attlog.drop_duplicates(subset=['NAME'], keep='last')
unique_ids = filter_unique_latest["ID"].astype(int).tolist()
unique_names = filter_unique_latest["NAME"].tolist()
departments = filter_unique_latest["DEPARTMENT"].tolist()
locations = filter_unique_latest["branchName"].tolist()

date_cols = df_attlog["DATE"].sort_values().unique().tolist()
df = pd.DataFrame(columns = ["EmpCode", "Name", "Department", "Location"] + date_cols)
df["EmpCode"] = unique_ids
df["Name"] = unique_names
df["Department"] = departments
df["Location"] = locations

for date in date_cols:
    daily_attendance = df_attlog[df_attlog["DATE"] == date]
    
    in_log = daily_attendance[daily_attendance["STATUS"]=="IN"]
    in_log = in_log.drop_duplicates(subset=['ID'], keep='first')
    
    out_log = daily_attendance[daily_attendance["STATUS"]=="OUT"]
    out_log = out_log.drop_duplicates(subset=['ID'], keep='last')
    
    for index1, row1 in df.iterrows():
        name = row1.Name
        for index_in, row_in in in_log.iterrows(): 
            in_time = row_in.TIME
            in_time = in_time.strftime("%H:%M")
            emp_name = row_in.NAME
            
            if emp_name == name:
                df.loc[index1, date] = str(in_time) + "-" + " "
        
    # dealing missing ins     
    df[date] = df[date].fillna("InMiss" + "-" + " ")
    
    for index1, row1 in df.iterrows():
        name = row1.Name
        for index_out, row_out in out_log.iterrows():
            out_time = row_out.TIME
            out_time = out_time.strftime("%H:%M")
            emp_name = row_out.NAME
            
            if emp_name == name:
                df.loc[index1, date] = str(df.loc[index1, date]).replace(" ", str(out_time))
    
    # dealing missing outs        
    df.loc[df[date].str.contains("- "), date] = df.loc[df[date].str.contains("- "), date].str.replace(" ", "OutMiss")

df = df.drop_duplicates(subset=["EmpCode"])  # dropping repetitive ids due to emp_name_spell changes over the month
df = df.fillna(" ")

#================================ TO EXCEL =========================================================================
# Autofit col widths 
root_dir = os.getcwd()
dump_path = root_dir + "\\InOut_ReportSAP.xlsx"
writer = pd.ExcelWriter(dump_path) 
df.to_excel(writer, sheet_name='DEC-InOutReport', index=False, na_rep='NaN')

for column in df:
    column_width = max(df[column].astype(str).map(len).max(), len(str(column)))
    col_idx = df.columns.get_loc(column)
    writer.sheets['DEC-InOutReport'].set_column(col_idx, col_idx, column_width)

writer.save()

# Color coded and col autofit
root_dir = os.getcwd()
dump_path = root_dir + "\\InOut_ReportSAP_Styled.xlsx" 
writer = pd.ExcelWriter(dump_path) 
df.style.applymap(
				lambda x: "background-color: red" if x=="InMiss-OutMiss" else "background-color: orange"  if "InMiss-" in str(x) else "background-color: orange"  if "-OutMiss" in str(x) else "background-color: green",
				subset = df.columns[4:]).to_excel(writer, sheet_name='DEC-InOutReport-Styled', index=False, na_rep='NaN')

for column in df:
    column_width = max(df[column].astype(str).map(len).max(), len(str(column)))
    col_idx = df.columns.get_loc(column)
    writer.sheets['DEC-InOutReport-Styled'].set_column(col_idx, col_idx, column_width)

writer.save()


# color mapped excel
# df.style.applymap(
# 				lambda x: "background-color: red" if x=="InMiss-OutMiss" else "background-color: orange"  if "InMiss-" in str(x) else "background-color: orange"  if "-OutMiss" in str(x) else "background-color: green",
# 				subset = df.columns[3:]).to_excel("style.xlsx")
