import pandas as pd
import openpyxl
from tkinter import filedialog
from decimal import Decimal

def process_excel_files(input_file_path, supporter_file_path, phone_file_path, output_file_path):
    
    snow_cases_df = pd.read_excel(input_file_path)
    chat_count_df = pd.read_excel(supporter_file_path)
    phone_count_df = pd.read_excel(phone_file_path)

    
    names_dict = {
 "A, Manjunatha": "Manjunatha A",
        "Kesavarapu, Sulochana":"Sulochana Kesavarapu",
        "Grand Total": "Grand Total"
    }

    
    phone_count_df['FullName'] = phone_count_df['FullName'].map(names_dict)

    
    result_df = snow_cases_df.merge(chat_count_df, how='left', left_on='Opened by', right_on='Supporter')
    result_df = result_df.merge(phone_count_df, how='left', left_on='Opened by', right_on='FullName')

    
    output_df = result_df[['Opened by', 'Chat', 'Total Chats Served', 'Phone', 'Incoming']].rename(columns={'Chat': 'chat created', 'Total Chats Served': 'chats taken', 'Phone': 'phones created', 'Incoming': 'phones taken'})

    
    output_df['Chat Percentage'] = ((output_df['chat created'] / output_df['chats taken']) * 100).round(2).astype(str) + '%'
    output_df['Chat Percentage'] = output_df['Chat Percentage'].apply(lambda x: '0' if x == 'nan%' else x)

    
    output_df['Phone Percentage'] = output_df.apply(lambda row: '0' if row['phones taken'] == 0 else str((Decimal(row['phones created'] / row['phones taken']) * 100).quantize(Decimal("0.00"))) + '%', axis=1)
    output_df['Phone Percentage'] = output_df['Phone Percentage'].apply(lambda x: '0' if x == 'nan%' else x)

    
    output_df['Phone Percentage'] = output_df['Phone Percentage'].replace('NaN%', '0')

    def color_format(val):
        if val == '0' or pd.isna(val):
            return 'background-color: white'  
        elif float(val.rstrip('%')) < 95.00:
            return 'background-color: red'
        else:
            return 'background-color: green'

    
    styled_table = output_df.style.applymap(color_format, subset=['Chat Percentage', 'Phone Percentage']).format({'Chat Percentage': "{:.2f}%", 'Phone Percentage': "{:.2f}%"})

    
    styled_table.to_excel(output_file_path, index=False)


input_file_path = "C:/Users/Cablet/Desktop/phone chat ratio/snow cases created.xlsx"
supporter_file_path = "C:/Users/Cablet/Desktop/phone chat ratio/chat count.xlsx"
phone_file_path = "C:/Users/Cablet/Desktop/phone chat ratio/phone count.xlsx"
output_file_path = "C:/Users/Cablet/Desktop/phone chat ratio/chatcountOutput.xlsx"


process_excel_files(input_file_path, supporter_file_path, phone_file_path, output_file_path)

print(f"Processing complete. Output saved to {output_file_path}")
