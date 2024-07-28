import pandas as pd
import openpyxl
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from statsmodels.tsa.seasonal import seasonal_decompose

df = pd.read_excel('SABCleaned_data.xlsx')

# Check for and handle negative Volume values if they are erroneous
df = df[df['Volume (hL)'] > 0]

# Data inspection
print(df.shape)
print(df.info())
print(df.head())
print(df['Volume (hL)'].describe())

# Visualize the distribution using box plots
plt.figure(figsize=(10, 6))
sns.boxplot(x=df['Volume (hL)'])
plt.title('Box Plot of Volume (hL)')
plt.show()

# Skewness and Kurtosis
print('Skewness: %f' % df['Volume (hL)'].skew())
print('Kurtosis: %f' % df['Volume (hL)'].kurt())

scaler = MinMaxScaler()
df['Volume (hL)_minmax'] = scaler.fit_transform(df[['Volume (hL)']])
print(df[['Volume (hL)', 'Volume (hL)_minmax']].head())

# Check the range of the scaled data
print("Description of Scaled Data:\n", df['Volume (hL)_minmax'].describe())

# Initialize the scaler
scaler = RobustScaler()

# Apply the scaler to the Volume (hL) column
df['Volume (hL)_robust'] = scaler.fit_transform(df[['Volume (hL)']])

# Check the first few rows to verify
print(df[['Volume (hL)', 'Volume (hL)_robust']].head())

# Datetime conversion
df['Date'] = pd.to_datetime(df['Year'], format='%Y')
df['Month'] = df['Date'].dt.month
df['Year'] = df['Date'].dt.year

# Set the date column as the index
df.set_index('Year', inplace=True)

# Group data for seasonal trends
monthly_sales = df.groupby(['Year', 'Month', 'Manufacturer'])['Volume (hL)_robust'].sum().reset_index()
print(monthly_sales)

# Plot the time series data
plt.figure(figsize=(14, 8))
plt.plot(df['Volume (hL)_robust'])
plt.xlabel('Year')
plt.ylabel('Volume (hL)')
plt.title('Time Series Plot of Volume')
plt.show()

# Plot the average monthly sales volume
plt.figure(figsize=(10, 6))
sns.lineplot(x='Month', y='Volume (hL)_robust', data=monthly_sales)
plt.title('Average Monthly Sales Volume')
plt.xlabel('Month')
plt.ylabel('Volume (hL)')
plt.show()

# Scatter plot of Seasonal Sales Volume against Year for different Manufacturers
plt.figure(figsize=(14, 8))
sns.scatterplot(data=df, x='Year', y='Volume (hL)_robust', hue='Manufacturer')
plt.title('Yearly Sales Volume by Manufacturer (Scaled)')
plt.xlabel('Year')
plt.ylabel('Scaled Volume (hL)')
plt.legend(title='Manufacturer')
plt.show()

# Boxplot of Volume by Pack Size
plt.figure(figsize=(14, 8))
sns.boxplot(data=df, x='Pack Size', y='Volume (hL)_robust')
plt.title('Boxplot of Scaled Volume (hL) by Pack Size')
plt.xlabel('Pack Size')
plt.ylabel('Scaled Volume (hL)')
plt.show()


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


