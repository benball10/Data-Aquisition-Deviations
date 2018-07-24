# Program that searches for deviations from normal data aquisition trends

import numpy as np
import pandas as pd
from statistics import mean

# Reading in table from Excel
# Excel table must have correct column titles (_______________)
df = pd.read_excel(r"C:\Users\Benjamin Ball\Documents\Python_Example.xlsx", index_col = None)

# Storing row and column lengths
row_count = df.shape[0]
col_count = df.shape[1]

# Adding new column of length row_count
new_array = np.array([0, 0])
new_array.resize(row_count, 1)

# Looping thru attribute column
j = df.Attribute[1]
for i in df.Attribute:
    if df.Attribute.iloc[i] != df.Attribute.iloc[i - 1]:
        arr_total = np.array()
        a = j
        while (a < i - 4):
            arr.append(df.Utilized_Total.iloc[a])
            a += 1        
        mean_total = mean(arr_total)
        
        arr_percent = np.array()
        b = j
        while (b < i - 4):
            arr.append(df.Utilized_Percent.iloc[b])
            b += 1        
        mean_percent = mean(arr_percent)
        
        #if (deviation from 7 days earlier) and (large enough data) -> compare to mean and stdev/mark if deviates
        #can make this test mark more to be safe (leave to manual review); 
        c = i - 4
        while (c < i - 1):
            if ((df.Utilized_Total.iloc[c] < .75 * df.Utilized_Total.iloc[c - 7]) and (df.Utilized_Total.iloc[c] > 1.25 * df.Utilized_Total.iloc[c - 7]) and (mean_total > 10000)):
                if ((df.Utilized_Total.iloc[c] < .75 * mean_total) or (df.Utilized_Total.iloc[c] > 1.25 * mean_total)):
                    new_array.iloc[c] = "Check this"

        d = i - 4
        while (d < i - 1):
            if ((df.Utilized_Percent.iloc[d] < .67 * df.Utilized_Percent.iloc[d - 7]) and (df.Utilized_Percent.iloc[d] > 1.33 * df.Utilized_Percent.iloc[d - 7])):
                if ((df.Utilized_Percent.iloc[d] < .67 * mean_percent) or (df.Utilized_Percent.iloc[d] > 1.33 * mean_percent)):
                    new_array.iloc[d] = "Check this"

        j = i

# Adding new_array to the end of the original table
df['Review'] = new_array

# Writing table to new Excel file
df.to_excel(r"C:\Users\Benjamin Ball\Documents\Python_Example2.xlsx", sheet_name = "sheet1", index = False)
