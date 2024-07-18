import pandas as pd


#Make dataframe from Excel sheet
df = pd.read_excel('Tech Stream Case Study_V3.xlsx')

column_to_convert = ['Volume (hL)']  # Column to convert

for column in column_to_convert:
    df[column] = df[column].apply(lambda x: float(x))#Convert to float(numeric)

# Keep only rows with a volume greater than 0
df = df[df['Volume (hL)'] > 0]

# Reset the index 
df.reset_index(drop=True, inplace=True)

# Save the cleaned dataset to a new Excel file
df.to_excel('SABCleaned_data.xlsx', index=False)

