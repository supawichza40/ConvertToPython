import pandas as pd
import numpy as np
import datetime as dt
#Name of the file you want to generate into.
month = "October2020Summary"
xls_payroll = pd.ExcelFile(r"C:\Users\supaw\Documents\DJ Payroll\Daily sheet October 2020.xlsx")#path of the original file.
day_1 = pd.read_excel(xls_payroll,"1")
day_temp_dt = pd.read_excel(xls_payroll,"{0}".format(4))

#column position row,column 1,2
#12,2
#24,2 = nan
#pd.isna(day_temp_dt.iloc[26][2])) = check for nan value

#hours location
#10,1
#22,1
#34,1
time = day_temp_dt.iloc[1][2]

print(time)