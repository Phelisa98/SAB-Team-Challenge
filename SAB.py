import pandas as pd
import openpyxl
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from statsmodels.tsa.seasonal import seasonal_decompose

df = pd.read_excel('Tech Stream Dataset_Amended_Cleaned.xlsx')

#Assumption_that_negativea_volume_or_=_0_are_erroneous

#Task_1

monthly_sales = df.groupby(['Region', 'Month'])['Volume (HL)'].sum().reset_index()
#Add all volume of and group by each region and month

plt.figure(figsize=(14, 8),dpi=500,facecolor='w',edgecolor='k')
for region in monthly_sales['Region'].unique():
    region_sales = monthly_sales[monthly_sales['Region'] == region]
    plt.plot(region_sales['Month'], region_sales['Volume (HL)'], label= region)


plt.title('Seasonal Sales Trends by Region')
plt.xlabel('Month')
plt.ylabel('Sales Volume')
plt.legend()
plt.show()

monthly_sales_price_tier = df.groupby(['Price Tier', 'Month','Region'])['Volume (HL)'].sum().reset_index()

#df.plot(x="City", y=["Population(million)", "Year(2020)"],
        #kind="line", figsize=(10, 10))

plt.figure(figsize=(14, 8))
for price_tier in monthly_sales_price_tier['Price Tier'].unique():
    price_tier_sales = monthly_sales_price_tier[monthly_sales_price_tier['Price Tier'] == price_tier]
    plt.plot(price_tier_sales['Month'], price_tier_sales['Volume'], label=price_tier)

plt.title('Seasonal Sales Trends by Price Tier')
plt.xlabel('Month')
plt.ylabel('Sales Volume')
plt.legend()
plt.show()

factors = df.groupby(['Month', 'Region', 'PackSize'])['Volume'].sum().reset_index()

# Visualize the impact of different factors using pair plots
sns.pairplot(factors, hue='Region', palette='Set1')
plt.show()

# Correlation matrix to see the relationships between factors
plt.figure(figsize=(10, 6))
sns.heatmap(factors.corr(), annot=True, cmap='coolwarm')
plt.title('Correlation Matrix of Factors Affecting Sales Volume')
plt.show()

plt.figure(figsize=(14, 8))
for price_tier in monthly_sales_price_tier['Price Tier'].unique():
    price_tier_sales = monthly_sales_price_tier[monthly_sales_price_tier['Price Tier'] == price_tier]
    plt.plot(price_tier_sales['Month'], price_tier_sales['Volume (HL)'], label=price_tier)

plt.title('Seasonal Sales Trends by Price Tier')
plt.xlabel('Month')
plt.ylabel('Sales Volume')
plt.legend()
plt.show()

factors = df.groupby(['Month', 'Region', 'Pack Size'])['Volume (HL)'].sum().reset_index()

# Visualize the impact of different factors using pair plots
sns.pairplot(factors, hue='Region', palette='Set1')
plt.show()

# Correlation matrix to see the relationships between factors
plt.figure(figsize=(10, 6))
sns.heatmap(factors.corr(), annot=True, cmap='coolwarm')
plt.title('Correlation Matrix of Factors Affecting Sales Volume')
plt.show()

