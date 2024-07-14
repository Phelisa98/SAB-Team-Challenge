import pandas as pd
import numpy as np
import openpyxl
import matplotlib.pyplot as plt

df = pd.read_excel('Tech Stream Dataset_Amended.xlsx')

df_cleaned = df[df['Volume (HL)'] > 0]

print(df_cleaned)


df_cleaned.to_excel('Tech Stream Dataset_Amended_Cleaned.xlsx' , index = False)
