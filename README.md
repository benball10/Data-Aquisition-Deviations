# Program that searches for deviations from normal data aquisition trends

import numpy as np
import pandas as pd
from statistics import mean

# Reading in table from Excel
# Excel table must have correct column titles (Attribute, Utilized_Total, and Utilized_Percent or modify code)
df = pd.read_excel(r"C:\Users\Ben\Documents\intentiq_example.xls", index_col = None)

# Storing row and column lengths
row_count = df.shape[0]
col_count = df.shape[1]

# Adding new column of length row_count
new_array = np.array([0, 0])
new_array.resize(row_count, 1)

# Looping through Attribute column
j = 1
for i in range(1, len(df.Attribute) - 1):
    if df.Attribute.iloc[i] != df.Attribute.iloc[i - 1]:
        
        if i - j < 4:  #change 4 to 10 or 11 bc compaing to 7 days ago?
            j = i

        else:
            # Calculating avg. total before recent 3 days
            arr_total = np.zeros((i - j - 4), dtype = float)
            a = j
        
            x = 0
            while (a < (i - 4)):
                np.put(arr_total, [x], df.Utilized_Total.iloc[a])  #check this! maybe a+1 or a+2
                a += 1
                x += 1

            if mean(arr_total) > 5000:
                n = len(arr_total) - 1
                for m in range(0, n):
                    if arr_total[m] == 0:
                        arr_total = np.delete(arr_total, (m))
                          
            print(arr_total)
            mean_total = mean(arr_total)
            print(mean_total)
                      
            # Calculating avg. percent before recent 3 days
            arr_percent = np.zeros((i - j - 4), dtype = float)
            b = j
            
            y = 0
            while (b < (i - 4)):
                np.put(arr_percent, [y], df.Utilized_Percent.iloc[b])
                b += 1
                y += 1
                
            if mean(arr_percent) > 1:
                p = len(arr_percent) - 1
                for o in range(0, p):
                    if arr_percent[o] == 0:
                        arr_percent = np.delete(arr_percent, (o))
                
            print(arr_percent)
            mean_percent = mean(arr_percent)
            print(mean_percent)

            #if (deviation from 7 days earlier) and (large enough data) -> compare to mean from prev. days/mark if deviates or "-"
            #can make this test mark more to be safe (leave to manual review); make size rule for percent
            c = i - 4
            while (c < i - 1):
                if ((df.Utilized_Total.iloc[c] < .75 * df.Utilized_Total.iloc[c - 7]) and (df.Utilized_Total.iloc[c] > 1.25 * df.Utilized_Total.iloc[c - 7]) and (mean_total > 10000)):
                    if ((df.Utilized_Total.iloc[c] < .75 * mean_total) or (df.Utilized_Total.iloc[c] > 1.25 * mean_total) or (df.Utilized_Total.iloc[c] == "-")):
                        new_array[c] = "Check this"

            d = i - 4
            while (d < i - 1):
                if ((df.Utilized_Percent.iloc[d] < .67 * df.Utilized_Percent.iloc[d - 7]) and (df.Utilized_Percent.iloc[d] > 1.33 * df.Utilized_Percent.iloc[d - 7])):
                    if ((df.Utilized_Percent.iloc[d] < .67 * mean_percent) or (df.Utilized_Percent.iloc[d] > 1.33 * mean_percent) or (df.Utilized_Total.iloc[c] == "-")):
                        new_array[d] = "Check this"
            
            j = i

# Adding new_array to the end of the original table
df['Review'] = new_array

# Writing table to new Excel file
df.to_excel(r"C:\Users\Ben\Documents\intentiq_example2.xlsx", sheet_name = "sheet1", index = False)
