import pandas as pd
import openpyxl
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from statsmodels.tsa.seasonal import seasonal_decompose

df = pd.read_excel('Tech Stream Dataset_Amended_Cleaned.xlsx')

#Assumption_that_negativea_volume_or_=_0_are_erroneous

#Task_1

monthly_sales = df.groupby(['Region', 'Month'])['Volume'].sum().reset_index()


