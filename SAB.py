import pandas as pd
import openpyxl
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from statsmodels.tsa.seasonal import seasonal_decompose

df = pd.read_excel('SABCleaned_data.xlsx')

# Check for and handle negative Volume values if they are erroneous
df = df[df['Volume (hL)'] >= 0]

# Data inspection
print(df.shape)
print(df.info())
print(df.head())

# Datetime conversion
df['Date'] = pd.to_datetime(df['Year'], format='%Y')
df['Month'] = df['Date'].dt.month
df['Year'] = df['Date'].dt.year

# Group data for seasonal trends
monthly_sales = df.groupby(['Year', 'Month', 'Manufacturer'])['Volume (hL)'].sum().reset_index()
print(monthly_sales)

#Plot seasonal sales trends by manufacturer
plt.figure(figsize=(14, 8))
sns.lineplot(data=monthly_sales,x='Month', y='Volume (hL)', hue='Manufacturer')
plt.title('Seasonal Sales Trends by Year')  
plt.xlabel('Month')
plt.ylabel('Volume (HL)')
plt.show()

#Group data for further pack type and pack size impact analysis
pack_sales = df.groupby(['Pack Type', 'Pack Size'])['Volume (hL)'].sum().reset_index()
print(pack_sales)

#Plot impact of Pack Type on Sales Volume
plt.figure(figsize=(14,8))
sns.barplot(data=pack_sales,x='Pack Type',y='Volume (hL)',ci=None)
plt.title('Impact of Pack Type on Sales Volume')
plt.xlabel('Pack Type')
plt.ylabel('Volume (hL)')
plt.show()

#Plot impact of Pack Size on Sales Volume
plt.figure(figsize=(14,8))
sns.barplot(data=pack_sales, x='Pack Size',y='Volume (hL)',ci=None)
plt.title('Impact of Pack Size on Sales Volume')
plt.xlabel('Pack Size')
plt.ylabel('Volume (hL)')
plt.show()

# Visualize impact of different factors using pair plots
sns.pairplot(pack_sales, palette='Set1')
plt.show()

#Correlation matrix to see the relationships between factors.
# First, we need to encode categorical variables to numeric format for correlation analysis
pack_sales_encoded = pd.get_dummies(pack_sales, columns=['Pack Type', 'Pack Size'])

#Add volume column back for correlation analysis
pack_sales_encoded['Volume (HL)'] = pack_sales['Volume (hL)']

# Correlation matrix to see the relationships between factors
plt.figure(figsize=(14, 8))
corr_matrix = pack_sales_encoded.corr()
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm',fmt='.2f', linewidths=0.5)
plt.title('Correlation Matrix of Factors Affecting Sales Volume')
plt.show()
