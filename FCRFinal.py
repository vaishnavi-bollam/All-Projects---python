import pandas as pd
import openpyxl
from openpyxl import load_workbook
import os
from tkinter import filedialog
from pathlib import Path
import re
import win32com.client
import os
import datetime

try:
    output_path = Path("C:/Users/Cablet/Desktop/FCR-csv")

    fcr_reports_folder_name = "FCR Reports"

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    inbox = outlook.GetDefaultFolder(6)

    desired_subject = "RE: FCR updated dump 1 hour"

    messages = inbox.Items

    for message in messages:
        subject = message.Subject

        
        if subject == desired_subject:
            attachments = message.Attachments

            current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            target_folder = output_path / fcr_reports_folder_name / current_time
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                filename = re.sub(r'[^0-9a-zA-Z.]+', '', attachment.FileName)
                attachment.SaveAsFile(target_folder / filename)
except Exception as e:
    print(f"Error processing Outlook messages: {e}")


from datetime import datetime
def get_most_recent_folder(folder_path):
    subfolders = [f.path for f in os.scandir(folder_path) if f.is_dir()]

    if not subfolders:
        print("No subfolders found in the specified path.")
        return None

    date_format = "%Y-%m-%d_%H-%M-%S"
    datetime_objects = [datetime.strptime(os.path.basename(subfolder), date_format) for subfolder in subfolders]

    most_recent_datetime = max(datetime_objects)
    most_recent_folder = most_recent_datetime.strftime(date_format)

    return os.path.join(folder_path, most_recent_folder)

file_path = "C:/Users/Cablet/Desktop/FCR-csv/FCR Reports"
most_recent_folder_path = get_most_recent_folder(file_path)

if most_recent_folder_path:
    print("Most recent folder:", most_recent_folder_path)
else:
    print("No valid folders found.")




import pandas as pd

column_mapping = {
    'number': 'Number',
    'sys_created_on' : 'Created',
    'contact_type': 'Source',
    'short_description' : 'Summary',
    'assignment_group': 'Assignment Group',
    'u_resolved': 'Resolved',
    'u_resolved_by': 'Assigned / Resolved',
    'sys_created_by': 'Created by',
    'u_reassignee_count': 'Assignee Count',
    'reassignment_count': 'Group Hop count',
    'reopen_count': 'Reopen count'
}

csv_file_path = most_recent_folder_path + "/sample.csv"


df = pd.read_csv(csv_file_path,encoding='latin1')

selected_columns = list(column_mapping.keys())
df_selected = df[selected_columns]

df_selected = df_selected.rename(columns=column_mapping)

csvto_excel_file_path = 'C:/Users/Cablet/Desktop/FCR-csv/FCRexceloutput.xlsx'

df_selected.to_excel(csvto_excel_file_path, index=False)

print(f"Conversion successful. Excel file saved at: {csvto_excel_file_path}")




#Load the input Excel file
input_file_path = csvto_excel_file_path
output_file_path = "C:/Users/Cablet/Desktop/FCRReport exe file/Dumpoutputfile.xlsx"
hcl_df = "C:/Users/Cablet/Desktop/FCRReport exe file/AgentDetails.xlsx"


df = pd.read_excel(input_file_path)
agent_details_df = pd.read_excel(hcl_df)

def FCRColumn(row):
    if (
        (row['Source'] == "Phone" and row['Assignee Count'] <= 1 and row['Group Hop count'] == 0 and row['Reopen count'] == 0 and row['Assignment Group'] == "GLB OFFICE SUPPORT FD") or
        (row['Source'] == "Chat" and row['Assignee Count'] <= 1 and row['Group Hop count'] == 0 and row['Reopen count'] == 0 and row['Assignment Group'] == "GLB OFFICE SUPPORT FD") or
        (row['Source'] == "Web" and row['Assignee Count'] >= 1 and row['Group Hop count'] == 0 and row['Reopen count'] == 0 and row['Assignment Group'] == "GLB OFFICE SUPPORT FD") or
        (row['Source'] == "Email" and row['Assignee Count'] >= 1 and row['Group Hop count'] == 0 and row['Reopen count'] == 0 and row['Assignment Group'] == "GLB OFFICE SUPPORT FD")
    ):
        return "Yes"
    else:
        return "No"
    

def calculate_site(created_by):
    match = agent_details_df[agent_details_df["ID"].str.strip().str.lower() == str(created_by).strip().lower()]

    if not match.empty:
        HCLQueue = match.iloc[0]["Country"]
        return HCLQueue.capitalize() 
    else:
        return "Not Available"


def calculate_resolved_by_country(assigned_to):
    match = agent_details_df[agent_details_df["Name"].str.strip().str.lower() == str(assigned_to).strip().lower()]

    if not match.empty:
        HCLQueue = match.iloc[0]["Country"]       
        return HCLQueue.capitalize() 
    else:
        return "Not Available"


df['FCR'] = df.apply(FCRColumn, axis=1)
df['Site'] = df['Created by'].apply(calculate_site)
df['Resolved by Country'] = df['Assigned / Resolved'].apply(calculate_resolved_by_country)

if os.path.exists(output_file_path):
    os.remove(output_file_path)

df.to_excel(output_file_path, index=False)
print("completed")


import pandas as pd
import numpy as np


pivot_output_file="C:/Users/Cablet/Desktop/FCRReport exe file/pivot.xlsx"


df1 = pd.read_excel(output_file_path)
grouped1 = df1.groupby(['Site', 'Source', 'FCR'])
count_data1 = grouped1.size().reset_index(name='Count')
pivot_table1 = count_data1.pivot_table(index=['Site', 'Source'], columns='FCR', values='Count', aggfunc='sum', fill_value=0, dropna=False)
pivot_table1.columns.name = None

pivot_table1.loc['Grand Total'] = pivot_table1.sum()
pivot_table1['Grand Total'] = pivot_table1.sum(axis=1)

print(pivot_table1)



# df2 = pd.read_excel(output_file_path)

# grouped2 = df2.groupby(['Source', 'FCR'])
# count_data2 = grouped2.size().reset_index(name='Count')
# pivot_table2 = count_data2.pivot_table(index=['Source'], columns='FCR', values='Count', aggfunc='sum', fill_value=0)

# pivot_table2.columns = ['Without FCR', 'Within FCR']

# pivot_table2.loc['Grand Total'] = pivot_table2.sum()
# pivot_table2['Grand Total'] = pivot_table2.sum(axis=1)

# pivot_table2['FCR%'] = (pivot_table2['Within FCR'] / pivot_table2['Grand Total']) * 100
# pivot_table2['FCR%'] = pivot_table2['FCR%'].round(2).astype(str) + '%'

# pivot_table2 = pivot_table2[['Without FCR', 'Within FCR', 'Grand Total', 'FCR%']]

# print(pivot_table2)


import pandas as pd
from datetime import datetime

# Read Excel file
df2 = pd.read_excel(output_file_path)

# Filter rows with today's date in the "Resolved" column
today = datetime.today().strftime('%m/%d/%Y')
filtered_df = df2[df2['Resolved'].dt.date == pd.to_datetime(today).date()]

# Group by ['Source', 'FCR'] after filtering
grouped2 = filtered_df.groupby(['Source', 'FCR'])

# Perform the same operations as before
count_data2 = grouped2.size().reset_index(name='Count')
pivot_table2 = count_data2.pivot_table(index=['Source'], columns='FCR', values='Count', aggfunc='sum', fill_value=0)

pivot_table2.columns = ['Without FCR', 'Within FCR']

pivot_table2.loc['Grand Total'] = pivot_table2.sum()
pivot_table2['Grand Total'] = pivot_table2.sum(axis=1)

pivot_table2['FCR%'] = (pivot_table2['Within FCR'] / pivot_table2['Grand Total']) * 100
pivot_table2['FCR%'] = pivot_table2['FCR%'].round(2).astype(str) + '%'

pivot_table2 = pivot_table2[['Without FCR', 'Within FCR', 'Grand Total', 'FCR%']]

print(pivot_table2)



import pandas as pd
df3 = pd.read_excel(output_file_path)
grouped3 = df3.groupby(['Site', 'Source', 'FCR'])
count_data3 = grouped3.size().reset_index(name='Count')
pivot_table3 = count_data3.pivot_table(index=['Site', 'Source'], columns='FCR', values='Count', aggfunc='sum', fill_value=0, dropna=False)

pivot_table3.columns = ['Without FCR', 'Within FCR']

pivot_table3.loc['Grand Total'] = pivot_table3.sum()
pivot_table3['Grand Total'] = pivot_table3.sum(axis=1)

pivot_table3['FCR%'] = (pivot_table3['Within FCR'] / pivot_table3['Grand Total']) * 100
pivot_table3['FCR%'] = pivot_table3['FCR%'].round(2).astype(str) + '%'

pivot_table3['FCR%'] = pivot_table3['FCR%'].apply(lambda x: '0' if x == 'nan%' else x)

pivot_table3 = pivot_table3[['Without FCR', 'Within FCR', 'Grand Total', 'FCR%']]

print(pivot_table3)


def color_format(val):
    if val == '0' or pd.isna(val):
        return ''
    elif float(val.rstrip('%')) < 61.43:
        return 'background-color: red'
    else:
        return 'background-color: green'


with pd.ExcelWriter(pivot_output_file, engine='openpyxl', mode='w') as writer:
    pivot_table1.to_excel(writer, sheet_name='PivotTable1', index=True)
    pivot_table2.to_excel(writer, sheet_name='PivotTable2', index=True)
    pivot_table3.to_excel(writer, sheet_name='PivotTable3', index=True)
    
    styled_table2 = pivot_table2.style.applymap(color_format, subset=['FCR%']).format({'FCR%': "{:.2f}%"})
    styled_table3 = pivot_table3.style.applymap(color_format, subset=['FCR%']).format({'FCR%': "{:.2f}%"})
    
    styled_table2.to_excel(writer, sheet_name='PivotTable2', index=True)
    styled_table3.to_excel(writer, sheet_name='PivotTable3', index=True)

print("pivot table generated")