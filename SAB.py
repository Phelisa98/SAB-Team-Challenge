import pandas as pd
import openpyxl
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from statsmodels.tsa.seasonal import seasonal_decompose

df = pd.read_excel('SABCleaned_data.xlsx')

# Check for and handle negative Volume values if they are erroneous
df = df[df['Volume (HL)'] >= 0]

# Data inspection
print(df.shape)
print(df.info())
print(df.head())

# Datetime conversion
df['Date'] = pd.to_datetime(df['Year'], format='%Y')
df['Month'] = df['Date'].dt.month
df['Year'] = df['Date'].dt.year

# Group data for seasonal trends
monthly_sales = df.groupby(['Year', 'Month', 'Manufacturer'])['Volume (HL)'].sum().reset_index()

# Plot seasonal sales trends by manufacturer
plt.figure(figsize=(14, 8), dpi=500, facecolor='w', edgecolor='k')
sns.lineplot(data=monthly_sales, x='Month', y='Volume (HL)', hue='Manufacturer')
plt.title('Seasonal Sales Trends by Manufacturer')
plt.xlabel('Month')
plt.ylabel('Sales Volume (HL)')
plt.legend()
plt.show()

# Group data by price tier and month
monthly_sales_price_tier = df.groupby(['Price Tier', 'Month', 'Manufacturer'])['Volume (HL)'].sum().reset_index()

# Plot seasonal sales trends by price tier
plt.figure(figsize=(14, 8))
for price_tier in monthly_sales_price_tier['Price Tier'].unique():
    price_tier_sales = monthly_sales_price_tier[monthly_sales_price_tier['Price Tier'] == price_tier]
    plt.plot(price_tier_sales['Month'], price_tier_sales['Volume (HL)'], label=price_tier)

plt.title('Seasonal Sales Trends by Price Tier')
plt.xlabel('Month')
plt.ylabel('Sales Volume (HL)')
plt.legend()
plt.show()

# Analyze impact of factors
factors = df.groupby(['Month', 'Manufacturer', 'Pack Size'])['Volume (HL)'].sum().reset_index()

# Visualize impact of different factors using pair plots
sns.pairplot(factors, hue='Manufacturer', palette='Set1')
plt.show()

factors = df.groupby(['Pack Type', 'Pack Size'])['Volume (HL)'].sum().reset_index()

# Visualize impact of different factors using pair plots
sns.pairplot(factors, palette='Set1')
plt.show()

# Correlation matrix to see the relationships between factors
plt.figure(figsize=(14, 8))
sns.heatmap(factors.corr(), annot=True, cmap='coolwarm')
plt.title('Correlation Matrix of Factors Affecting Sales Volume')
plt.show()
