import pandas as pd
from settings import *

# Settings


# Read data from excels
df = pd.read_excel('report.xlsx')

# Extract important attributes
df1 = df[['Friendly_Name','Problem_ID', 'Message', 'Severity']]

# Count problems and create a new data frame
df1 = df1.fillna(value='')
if romania_only:
    df1 = df1[df1['Friendly_Name'].str.contains("RO-")]
problems_index = (df1['Problem_ID'].value_counts()).to_frame()
problems_index.rename(columns = {'Problem_ID' : 'NO_Problems'}, inplace=True)

# Ignore problems
for name in problem_ignore:
    problems_index = problems_index.drop(name)

entries = problems_index.size

# Set total of problems
problems_index.loc['Total'] = problems_index.sum(axis=0)

# Get problem description
tag = pd.read_excel('problem_description.xlsx', index_col='Name')
problems = problems_index.reset_index()
problems.rename(columns = {'index' : 'Problem_ID'}, inplace=True)

list_prob = (problems['Problem_ID'].to_numpy()).tolist()
tag = tag.loc[list_prob]
no_problems = problems['NO_Problems']
problems_index['Explanation'] = tag['Explanation']


# Create a Pandas Excel writer using XlsxWriter as the engine and write problems to excel.
writer = pd.ExcelWriter('Data.xlsx', engine='xlsxwriter')
problems_index.to_excel(writer, sheet_name='Sheet1')

# Access the XlsxWriter workbook and worksheet objects from the dataframe.
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Table formatting
critical = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'center'})

for prob in problems['Problem_ID']:
    if prob != 'Total':
        [sev] = df1.loc[df1['Problem_ID'] == prob].head(1)['Severity']
        if sev == 'C':
            [index] = problems.index[problems['Problem_ID'] == prob]
            worksheet.set_row_pixels(index + 1, 20, critical)
            
center = workbook.add_format({'align': 'center'})
worksheet.set_column(2, 2, 60, center)
worksheet.set_column(1, 1, 15, center)


# Create a chart objects.
chart1 = workbook.add_chart({'type': 'pie'})
chart2 = workbook.add_chart({'type': 'column'})


# Configure chart1
chart1.set_title({'name': 'Number of problems'})
chart1.add_series({
    'categories': f'=Sheet1!A2:A{entries}',
    'values':     f'=Sheet1!B2:B{entries}',
})

# Configure chart2
chart2.set_title({'name': 'Number of problems'})
chart2.set_legend({'none': True})
chart2.add_series({
    'categories': f'=Sheet1!A2:A{entries}',
    'values':     f'=Sheet1!B2:B{entries}',
    'gap':    10,
})

# Insert the chart into the worksheet.
worksheet.insert_chart('D1', chart1)
worksheet.insert_chart('D17', chart2)

# Close the Pandas Excel writer and output the Excel file.
writer.save()


   