import pandas as pd
import datetime as dt
import win32com.client
def checkNameAndAppendToArray(name,array_hour,array_money,day_temp_dt):

    if (str(day_temp_dt.iloc[1][2]).lower() == name.lower()):
        time = day_temp_dt.iloc[10, 1]
        array_hour.append(time)
        h, m, s = str(time).split(":")
        array_money.append(((int(h) * 20) + ((int(m) / 60) * 20)))
    elif (str(day_temp_dt.iloc[12][2]).lower() == name.lower()):
        time = day_temp_dt.iloc[22, 1]
        array_hour.append(time)
        h, m, s = str(time).split(":")
        array_money.append(((int(h) * 20) + ((int(m) / 60) * 20)))

    elif (str(day_temp_dt.iloc[24][2]).lower() == name.lower()):
        time = day_temp_dt.iloc[34, 1]
        array_hour.append(time)
        h, m, s = str(time).split(":")
        array_money.append(((int(h) * 20) + ((int(m) / 60) * 20)))

    else:
        array_hour.append(dt.timedelta(hours=0, minutes=0, seconds=0))
        array_money.append(0)

#Name of the file you want to generate into.
#STEP0: USE THE FILE FORMAT GIVEN, OTHERWISE WILL NOT WORK.
#STEP1: CHOOSE THE NAME OF THE FILE YOU WANT TO CALL
month = "October2020Summary"
#STEP2: FIND THE PATH WHERE THE FILE IS
xls_payroll = pd.ExcelFile(r"C:\Users\supaw\Documents\DJ Payroll\Daily sheet October 2020.xlsx")#path of the original file.

day_1 = pd.read_excel(xls_payroll,"1")
daily_card_array = [0]
daily_cash_array = [0]
daily_total_array = [0]
day = [0]
day_counter = 1
#STEP3: CREATE A NEW STAFF VARIABLE
#Add staff list with name, hours and money
joy_hours_array = [0]
joy_money_array = [0]

si_hours_array = [0]
si_money_array = [0]

ann_hours_array=[0]
ann_money_array=[0]

na_hours_array=[0]
na_money_array=[0]
for i in range(31):
    day_temp_dt = pd.read_excel(xls_payroll,"{0}".format(day_counter))
    day.append(day_counter)
    if(day_counter==4):
        print(4)
    day_counter+=1
    daily_card_array.append(day_temp_dt.iloc[47][2])
    daily_cash_array.append(day_temp_dt.iloc[47][1])
    daily_total_array.append(day_temp_dt.iloc[47][1]+day_temp_dt.iloc[47][2])


    checkNameAndAppendToArray("joy",joy_hours_array,joy_money_array,day_temp_dt)
    checkNameAndAppendToArray("NAN", na_hours_array, na_money_array, day_temp_dt)#NAN IS USE INSTEAD OF NA, SINCE NA = NOT APPLICABLE AND CAUSE IT TO HAVE VALUE NAN.
    checkNameAndAppendToArray("si", si_hours_array, si_money_array, day_temp_dt)
    checkNameAndAppendToArray("ann", ann_hours_array, ann_money_array, day_temp_dt)
    #STEP4:ADD THE NEW PERSON INTO THE FUNCTION, AND FOLLOW THE ABOVE EXAMPLE. E.G.checkNameAndAppendToArray("ann", ann_hours_array, ann_money_array, day_temp_dt)







day.append("Total")
daily_cash_array.append(sum(daily_cash_array,1))
daily_card_array.append(sum(daily_card_array))
daily_total_array.append(sum(daily_total_array))
day_df = pd.DataFrame(day,columns=["Days"])

cash_df = pd.DataFrame(daily_cash_array,columns=["Cash"])




card_df = pd.DataFrame(daily_card_array,columns=["Card"])
total_df = pd.DataFrame(daily_total_array,columns=["Total"])
#create a staff list dataframe
#STEP5:CREATE A VARIABLE DATAFRAME TO BE ADDED.
joy_hour_df = pd.DataFrame(joy_hours_array,columns=["Joy Hour"])
joy_money_df =  pd.DataFrame(joy_money_array,columns=["Joy Money"])
si_hour_df = pd.DataFrame(si_hours_array,columns=["Si Hour"])
si_money_df =  pd.DataFrame(si_money_array,columns=["Si Money"])
ann_hour_df = pd.DataFrame(ann_hours_array,columns=["Ann Hour"])
ann_money_df =  pd.DataFrame(ann_money_array,columns=["Ann Money"])
na_hour_df = pd.DataFrame(na_hours_array,columns=["Na Hour"])
na_money_df =  pd.DataFrame(na_money_array,columns=["Na Money"])


merge_df = pd.concat([day_df,cash_df,card_df,total_df,joy_hour_df,joy_money_df,si_hour_df,si_money_df,ann_hour_df,ann_money_df,na_hour_df,na_money_df],axis=1,join="outer")
#STEP6: ADD THE DATAFRAME, TO THE END OF THE LIST OF DF TO CONCATENATE.
print(merge_df)
merge_df = merge_df.drop(labels = 0,axis = 0)
merge_df.to_excel(month+".xlsx",index=False)

print("Finish summarise, this month gross profit.")
#STEP7: FINISH GENERATE END OF MONTH STATEMENT SUMMARY.

