import pandas as pd
from settings import *

# Settings


# Read data from excels
df = pd.read_excel('report.xlsx')

# Extract important attributes
df1 = df[['Friendly_Name','Problem_ID', 'Message']]

# Count problems and create a new data frame
df1 = df1.fillna(value='')
if romania_only:
    df1 = df1[df1['Friendly_Name'].str.contains("RO-")]
problems_index = (df1['Problem_ID'].value_counts()).to_frame()
problems_index.rename(columns = {'Problem_ID' : 'NO_Problems'}, inplace=True)

# Ignore problems
for name in problem_ignore:
    problems_index = problems_index.drop(name)
    
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Data.xlsx', engine='xlsxwriter')
problems_index.to_excel(writer, sheet_name='Sheet1')

# Access the XlsxWriter workbook and worksheet objects from the dataframe.
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Create a chart objects.
chart1 = workbook.add_chart({'type': 'pie'})
chart2 = workbook.add_chart({'type': 'column'})


# Configure chart1
chart1.set_title({'name': 'Number of problems'})
chart1.add_series({
    'categories': '=Sheet1!A2:A36',
    'values':     '=Sheet1!B2:B36',
})

# Configure chart2
chart2.set_title({'name': 'Number of problems'})
chart2.set_legend({'none': True})
chart2.add_series({
    'categories': '=Sheet1!A2:A36',
    'values':     '=Sheet1!B2:B36',
    'gap':    10,
})

# Insert the chart into the worksheet.
worksheet.insert_chart('D1', chart1)
worksheet.insert_chart('D17', chart2)

# Close the Pandas Excel writer and output the Excel file.
writer.save()