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

input_file_name = 'inputfile'  

csv_file_path = f'C:/Users/Cablet/Desktop/DailyTicketshandled/{input_file_name}.csv'
excel_file_path = f'C:/Users/Cablet/Desktop/DailyTicketshandled/{input_file_name}.xlsx'

if os.path.exists(csv_file_path):
    input_file_path = csv_file_path
elif os.path.exists(excel_file_path):
    input_file_path = excel_file_path
else:
    raise FileNotFoundError(f"Neither {input_file_name}.csv nor {input_file_name}.xlsx found in the specified directory.")

df = read_input_file(input_file_path)


task_counts = df.groupby('User')['Task'].count().reset_index()

duplicate_task_counts = df.groupby('User')['Task'].nunique().reset_index()

df['Date'] = pd.to_datetime(df['Updated']).dt.date

datewise_task_counts = df.groupby(['User', 'Date']).size().reset_index(name='Count of Task')

total_tasks_count = datewise_task_counts.groupby('User')['Count of Task'].sum().reset_index(name='Total Count')

result_df = pd.DataFrame(columns=['User', 'Count', 'Count of Task'])
for _, row in total_tasks_count.iterrows():
    
    user_datewise_counts = datewise_task_counts[datewise_task_counts['User'] == row['User']]
    
    result_df = pd.concat([result_df, pd.DataFrame({'User': [row['User']], 'Count': [row['Total Count']], 'Count of Task': ['']}), 
                          user_datewise_counts[['User', 'Date', 'Count of Task']]], ignore_index=True)

output_file_path = 'C:/Users/Cablet/Desktop/DailyTicketshandled/DailyTicketshandledOutput.xlsx'

def color_format(val):
    if pd.notna(val):
        if val >= 30:
            return 'background-color: green'
        elif 20 <= val < 30:
            return 'background-color: orange'
        else:
            return 'background-color: red'
    else:
        return ''

with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    task_counts.to_excel(writer, sheet_name='with_duplicate_count', index=False)
    duplicate_task_counts.to_excel(writer, sheet_name='without_duplicate_count', index=False)
    result_df.style.applymap(color_format, subset=['Count']).to_excel(writer, sheet_name='combined_task_count', index=False)


print(f"Task counts saved to {output_file_path}")
