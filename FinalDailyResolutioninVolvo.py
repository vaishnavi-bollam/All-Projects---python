import pandas as pd
import os

def read_input_file(file_path):
    _, file_extension = os.path.splitext(file_path)
    
    if file_extension.lower() == '.xlsx':
        # Read Excel file
        return pd.read_excel(file_path)
    elif file_extension.lower() == '.csv':
        # Read CSV file
        return pd.read_csv(file_path, encoding='latin1')
    else:
        raise ValueError("Unsupported file format. Please provide an Excel (.xlsx) or CSV (.csv) file.")

input_file_name = 'input'  

csv_file_path = f'C:/Users/a221616/Desktop/Dailyresolution/{input_file_name}.csv'
excel_file_path = f'C:/Users/a221616/Desktop/Dailyresolution/{input_file_name}.xlsx'

if os.path.exists(csv_file_path):
    input_file_path = csv_file_path
elif os.path.exists(excel_file_path):
    input_file_path = excel_file_path
else:
    raise FileNotFoundError(f"Neither {input_file_name}.csv nor {input_file_name}.xlsx found in the specified directory.")

df = read_input_file(input_file_path)

task_counts = df.groupby('Assigned To')['Number'].count().reset_index()
duplicate_task_counts = df.groupby('Assigned To')['Number'].nunique().reset_index()

df['Date'] = pd.to_datetime(df['Resolved']).dt.date
datewise_task_counts = df.groupby(['Assigned To', 'Date']).size().reset_index(name='Count of Task')
total_tasks_count = datewise_task_counts.groupby('Assigned To')['Count of Task'].sum().reset_index(name='Total Count')

result_df = pd.DataFrame(columns=['Assigned To', 'Count', 'Count of Task'])
for _, row in total_tasks_count.iterrows():
    user_datewise_counts = datewise_task_counts[datewise_task_counts['Assigned To'] == row['Assigned To']]
    result_df = pd.concat([result_df, pd.DataFrame({'Assigned To': [row['Assigned To']], 'Count': [row['Total Count']], 'Count of Task': ['']}), 
                           user_datewise_counts[['Assigned To', 'Date', 'Count of Task']]], ignore_index=True)

output_file_path = 'C:/Users/a221616/Desktop/Dailyresolution/dailyresolutionOutput.xlsx'


def color_format(val):
    if pd.notna(val):
        if val >= 15:
            return 'background-color: green'
        elif 12 <= val < 15:
            return 'background-color: grey'
        else:
            return 'background-color: red'
    else:
        return '' 

with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    result_df.style.applymap(color_format, subset=['Count']).to_excel(writer, sheet_name='combined_task_count', index=False)

print(f"Task counts saved to {output_file_path}")
