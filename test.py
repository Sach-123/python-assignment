import pandas as pd


# note:- it is also possible to save both the input sheet in same excel file,
# but i have decided to make seperate file
pd.set_option('display.max_columns',None)

# reading the two sheets into data frames
sheet1 = pd.read_excel('User IDs.xlsx')
sheet2 = pd.read_excel('Rigorbuilder RAW.xlsx')

#correcting the mistake in team name 'Brandtech Lab' in input sheet 1
sheet1['Team Name'] = sheet1['Team Name'].replace('Brandtech Lab', 'BrandTech Lab', regex=True)


# merge the two data frames based on 'User ID' and 'uid'
merged = pd.merge(sheet1, sheet2, left_on='User ID', right_on='uid')


# group the merged data frame by 'Team Name' and calculate the mean values
grouped = merged.groupby('Team Name').mean()[['total_statements', 'total_reasons']]
grouped['Average Statements'] = grouped['total_statements']
grouped['Average Reasons'] = grouped['total_reasons']
grouped = grouped.drop(['total_statements', 'total_reasons'], axis=1)

# sort the resulting data frame by the average statements and print it
ranked = grouped.sort_values(by=['Average Statements'], ascending=False)
ranked = ranked.reset_index()
ranked.index += 1

# renaming the 1st two column name according to the given output sheet
ranked = ranked.reset_index().rename(columns={'index': 'Team Rank', 'Team Name': 'Thinking Teams Leaderboard'})

ranked['Average Statements'] = ranked['Average Statements'].apply(lambda x: round(x, 2)).astype(float)
ranked['Average Reasons'] = ranked['Average Reasons'].apply(lambda x: round(x, 2)).astype(float)

print(ranked)

# creating output sheet 1
writer = pd.ExcelWriter('Leaderboard TeamWise (Output).xlsx', engine='openpyxl')
ranked.to_excel(writer, index=False, sheet_name='Sheet1' , float_format="%.2f")

workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Set the default number format for all cells to two decimal places
number_format = '#0.00'
for col in range(1, worksheet.max_column + 1):
    for row in range(1, worksheet.max_row + 1):
        cell = worksheet.cell(row=row, column=col)
        if isinstance(cell.value, float):
            cell.number_format = number_format
writer.close()

#above I have created output sheet 1 

# now working to create output sheet 2
data = pd.read_excel('Rigorbuilder RAW.xlsx')

# create a new column for the sum of total_statements and total_reasons
data['sum'] = data['total_statements'] + data['total_reasons']

# below line is writen to consider alphabet before special character 
data['lower_name'] = data['name'].str.replace('[^a-zA-Z0-9\s]', '').str.lower()

# sort the data by the sum and name columns
data = data.sort_values(by=['sum', 'lower_name'], ascending=[False, True])

#after data is sortef we can drop the new column we created
data.drop('lower_name', axis=1, inplace=True)

# add a new column for the Rank
data['Rank'] = range(1, len(data)+1)

# selecting, reordering and renaming the desired columns
output = data[['Rank', 'name', 'uid', 'total_statements', 'total_reasons']]
output = output.rename(columns={'name': 'Name','uid':'UID','total_statements':'No. of Statements','total_reasons':'No. of Reasons'})

# print the output
print(output)

# creating output sheet 2
writer = pd.ExcelWriter('Leaderboard Individual (Output).xlsx', engine='openpyxl')
output.to_excel(writer, index=False, sheet_name='Sheet1')
writer.close()

# note: i have created two different sheet for two different output for better readability
# also the code require 'openpyxl' module to be installed in computer
# one more thing in "Leaderboard TeamWise (Output)" there is a mistake in second cell of "Average Reason" it should be 10 but you have given 9.33.