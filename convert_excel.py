import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import pandas as pd

def ErrorBox(message):
    messagebox.showerror('Error', message)

BASE_DIR = Path(__file__).resolve().parent



root = tk.Tk()
input_filename = filedialog.askopenfilename(
    initialdir=BASE_DIR,
    title='Select an Excel File to READ',
    filetypes=[("Excel Files", "*.xlsx")]
)

output_filename = filedialog.askopenfilename(
    initialdir=BASE_DIR,
    title='Select NORMALIZED_TABLE excel file, press CANCEL if none',
    filetypes=[("Excel Files", "*.xlsx")]
)

root.destroy()

# test if correct input file is selected
if input_filename == '':
    ErrorBox('Select an Input Excel File')
    
if input_filename == output_filename:
    ErrorBox('Input and Output File must not be the same')
    
input_data = pd.ExcelFile(input_filename)

if 'PROJECT' not in input_data.sheet_names:
    ErrorBox('No PROJECT Sheet found')
    
# create table for project
try:
    project = input_data.parse('PROJECT', dtype={'DONOR_CODE': object}).copy()
    project_table = project[[
        'PROJECT_TITLE',
        'PROJECT_ID',
        'PROJECT_DESCRIPTION',
        'TOTAL_REQUIRED_RESOURCE',
        'TOTAL_AVAILABLE_RESOURCE',
        'COMMENT',
        'APPROVAL_DATE'
    ]]
    # table for PROJECT sheet
    project_table = project_table.dropna(subset=['PROJECT_ID'])
    # Project ID
    this_project_id = project_table.iloc[0]['PROJECT_ID']
    
    # table for DONOR sheet
    donor_table = project[['DONOR_NAME', 'DONOR_CODE']]
    donor_table = donor_table.dropna()
    
    # table for PROJECT_RESOURCES sheet
    project_resources_table = project[['RESOURCE_TYPE', 'DONOR_CODE', 'AMOUNT']]
    project_resources_table = project_resources_table.dropna(subset=['RESOURCE_TYPE'])
    project_resources_table['PROJECT_ID'] = this_project_id
    
    # table for PROJECT_POST_TITLE sheet
    project_post_title_table = project[['CLEARANCE_ROLE', 'NAME']]
    project_post_title_table = project_post_title_table.dropna()
    project_post_title_table = project_post_title_table.rename(columns={'CLEARANCE_ROLE':'POST_TITLE'})
    project_post_title_table['PROJECT_ID'] = this_project_id
    
    # table for PROJECT_APPROVAL sheet
    project_approval_table = project[['CLEARANCE_ROLE','NAME','APPROVAL_DATE']]
    project_approval_table = project_approval_table.dropna(subset=['CLEARANCE_ROLE'])
    project_approval_table['PROJECT_ID'] = this_project_id
    
    # columns for the other sheets
    budget_code_columns = ['CODE', 'DESCRIPTION']
    activity_budget_columns = ['PLANNED_ACTIVITY', 'BUDGET_CODE']
    output_columns = ['PROJECT_ID', 'PROJECT_OUTCOME', 'PROJECT_OUTPUT']
    indicator_columns = [
        'PROJECT_OUTPUT',
        'INDICATOR',
        'DISAGGREGATION',
        'BASELINES',
        'YEAR',
        'ANNUAL_TARGET',
        'UNIT',
        'MID_YEAR_ACTUALS',
        'END_YEAR_ACTUALS',
        'DATA_SOURCE'
    ]
    
    activity_columns = [
        'PROJECT_OUTPUT',
        'PLANNED_ACTIVITY',
        'Q1',
        'Q2',
        'Q3',
        'Q4',
        'RESPONSIBLE_PARTY',
        'DONOR_CODE',
        'ANNUAL_BUDGET',
        'AMOUNT_FUNDED',
        'AMOUNT_UNFUNDED',
        'Q1_EXPENDITURE',
        'Q2_EXPENDITURE',
        'Q3_EXPENDITURE',
        'Q4_EXPENDITURE',
        'PROGRESS'
    ]
    activity_country_columns = ['PLANNED_ACTIVITY','COUNTRY']
    
    # tables for other sheets
    budget_code_table = pd.DataFrame(columns=budget_code_columns)
    activity_budget_table = pd.DataFrame(columns=activity_budget_columns)
    output_table = pd.DataFrame(columns=output_columns)
    indicator_table = pd.DataFrame(columns=indicator_columns)
    activity_table = pd.DataFrame(columns=activity_columns)
    activity_country_table = pd.DataFrame(columns=activity_country_columns)
    
    count = 1
    # for each output and indicator sheets insert data
    while True:
        try:
            activity = input_data.parse(f'OUTPUT {count}', header=2)
            
            # adding to the output table
            output = input_data.parse(f'OUTPUT {count}').iloc[0:1]
            output['PROJECT_ID'] = this_project_id
            output = output[output_columns]
            output_table = pd.concat([output_table, output], ignore_index=True)
            output_table = output_table.drop_duplicates()
            
            for index, row in activity.iterrows():
                #add to budget_code and activity_budget table
                if not pd.isna(row['BUDGET_DESCRIPTION']):
                    for budget_description in row['BUDGET_DESCRIPTION'].split(','):
                        code = budget_description.split('-')[0].strip()
                        description = budget_description.split('-')[1].strip().lower().capitalize()
                        
                        budget_code = pd.DataFrame([[code, description]], columns=budget_code_columns)
                        budget_code_table = pd.concat([budget_code_table, budget_code], ignore_index=True)
                        
                        activity_budget = pd.DataFrame([[row['PLANNED_ACTIVITY'], code]], columns=activity_budget_columns)
                        activity_budget_table = pd.concat([activity_budget_table, activity_budget], ignore_index=True)
                else:
                    activity_budget = pd.DataFrame([[row['PLANNED_ACTIVITY'], row['BUDGET_DESCRIPTION']]], columns=activity_budget_columns)
                    activity_budget_table = pd.concat([activity_budget_table, activity_budget], ignore_index=True)
                    
                activity_budget_table = activity_budget_table.drop_duplicates()
                
                # add to activity_country table
                
                for country in row['COUNTRY'].split(','):
                    country_name = country.lower().capitalize()
                    country_df = pd.DataFrame([[row['PLANNED_ACTIVITY'], country_name]], columns=activity_country_columns)
                    activity_country_table = pd.concat([activity_country_table, country_df], ignore_index=True)
                    
                activity_country_table = activity_country_table.drop_duplicates()
                
            activity['PROJECT_OUTPUT'] = output.iloc[0]['PROJECT_OUTPUT']
            new_df = activity[activity_columns]
            activity_table = pd.concat([activity_table, new_df], ignore_index=True)
            activity_table = activity_table.drop_duplicates()
                
        except ValueError:
            break
        
        try:
            # adding to indicator table
            indicator = input_data.parse(f'INDICATOR {count}')
        except ValueError:
            pass
        
        
        count += 1
    
except Exception as e:
    ErrorBox(e)

    
# test if correct output file is selected
output_sheets = [
    'PROJECT',
    'DONOR',
    'PROJECT_RESOURCE',
    'PROJECT_POST_TITLE',
    'PROJECT_APPROVAL',
    'OUTPUT',
    'PLANNED_ACTIVITY',
    'ACTIVITY_COUNTRY',
    'BUDGET_CODE',
    'ACTIVITY_BUDGET',
    'INDICATOR'
]

if output_filename != '':
    output_data = pd.ExcelFile(output_filename)
    if output_sheets != output_data.sheet_names:
        ErrorBox('Incorrect NORMALIZED_TABLE file selected. Click CANCEL if None')