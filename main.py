import pandas as pd
import numpy as np
import os
from datetime import datetime
import calendar

STATEMENTS_DIRECTORY = 'statements/'
HISTORY_FILE = 'history.txt'
EXCEL_FILE = 'bank.xlsx'

def read_history():
    mode = 'r' if os.path.exists(HISTORY_FILE) else 'w+'
    f = open(HISTORY_FILE, mode)
    history = [line.strip() for line in f.readlines()]
    f.close()
    return history

def process_statement(statement):
    df = pd.read_csv(f'{STATEMENTS_DIRECTORY}/{statement}')
    df = df.fillna('')
    latest_date = df['Post Date'].values[0]
    dt = datetime.strptime(latest_date, '%m/%d/%Y')
    statement_month = f'{calendar.month_name[dt.month]}_{dt.year}'
    df = df[df['Category'] != '']
    
    # actually processing
    g = df.groupby('Category').sum().abs()

    with pd.ExcelWriter(EXCEL_FILE) as writer:
        raw_sheet = f'{statement_month}_RAW'
        summary_sheet = f'{statement_month}_SUMMARY'
        
        df.to_excel(writer, sheet_name=raw_sheet)
        g.to_excel(writer, sheet_name=summary_sheet)
        total_spent = sum(g['Amount'])

        
        # generate chart
        max_row = len(g) + 1
        workbook = writer.book
        worksheet = writer.sheets[summary_sheet]
        chart = workbook.add_chart({'type': 'pie'})
        chart.add_series({
            'categories': f'={summary_sheet}!A2:A{max_row}',
            'values': f'={summary_sheet}!B2:B{max_row}'
        })    
        worksheet.insert_chart('D2', chart)    
        
        # add total amount spent
        max_row_padded = max_row + 2
        bold = workbook.add_format({'bold': True, 'align': 'center'})
        worksheet.write(f'A{max_row_padded}', 'Total', bold)
        worksheet.write(f'B{max_row_padded}', total_spent)
        
        
def read_new_statements():
    history = read_history()
    new_statements = [x.strip() for x in os.listdir(STATEMENTS_DIRECTORY) if x.strip() not in history]
    return new_statements

def process_statements(statements):
    f = open(HISTORY_FILE, 'a')
    for statement in statements:
        process_statement(statement)
        f.write(f'{statement}\n')
    f.close()
    
def clean_up():
    f = open(HISTORY_FILE, 'r')
    f.close()
    os.remove(HISTORY_FILE)
    os.remove(EXCEL_FILE)
    

def main():
    print('READING IN NEW STATEMENTS...')
    statements = read_new_statements()
    count_new = len(statements)
    print(f'PROCESSING {count_new} STATEMENTS...')
    process_statements(statements)
    print('FINISHED PROCESSING STATEMENTS...OPENING FILE...')
    os.startfile(EXCEL_FILE)
    # for test
    # clean_up()

if __name__ == '__main__':
    main()