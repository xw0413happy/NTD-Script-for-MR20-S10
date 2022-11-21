#!/usr/bin/env python
# coding: utf-8

# # Scheduled Service and Actual Service

# Title: Calculating actual vehicle revenue miles & hours
# Contact: wxi@leegov.com
# Author: Dr. Liang Zhai, Wang Xi
# Last Updated: 11-18-2022

# File Format: xlsx format

# Update notes:
# #### Schedule Tables (Input):
# - Daily Ridership by Route
# - Service changes within the RY/FY (reporting year / fiscal year)
# - Schduled VRM & VRH (vehicle revenue miles & vehicle revenue hours)
# - Schduled DHM & DHH (deadhead miles & deadhead hours)
# 
# #### Deviation Tables (Input):
# - Atypical days
# - Added Runs
# - Lost Runs
# 
# #### List of Routes (Input)
# - Any route used during the RY/FY 
# - Update the list on line 19 in the script below 
# 
# #### Calculated Actual Service Tables (Output):
# - Actual VRM & VRH (actual vehicle revenue miles & actual vehicle revenue hours)
# - Actual TVM & TVR (actual total vehicle miles & actual total vehicle hours)

# In[73]:

# Import libraries
import os
import pandas as pd
import numpy as np
import datetime
from datetime import date, timedelta

# Set working directory
os.chdir(r'S:\LeeTran\Planning\National Transit Database\Archived Service Data\Annual Report Packages\2021-2030\RY 2022\Forms\S-10 MB DO\NTD Script tool FY22')
os.getcwd()

print('* Importing input tables...')
df_ridership = pd.read_excel('1_Daily Ridership by Route.xlsx') # Daily Ridership by Route
df_service_change = pd.read_excel('2_Service Changes.xlsx') # Service Change table
df_sched_mile_manual = pd.read_excel('3_Scheduled Miles.xlsx') # Scheduled Miles
df_sched_hour_manual = pd.read_excel('4_Scheduled Hours.xlsx') # Scheduled Hour
df_DHM = pd.read_excel('5_Deadhead Vehicle Miles.xlsx') # deadhead miles
df_DHH = pd.read_excel('6_Deadhead Vehicle Hours.xlsx') # deadhead hours
df_atypical = pd.read_excel('7_Atypical Days.xlsx') # atypical days
df_added_run = pd.read_excel('8_Added Runs.xlsx') # Added Runs
df_lost_run = pd.read_excel('9_Lost Runs.xlsx') # Lost Runs
print('* Input tables imported.')


# route list
ls_route = []
 
# iterating the columns
for col in df_ridership.columns:
    ls_route.append(col)
# print(ls_route_str)

remov_ls = ['Service Type', 'Month', 'Date', 'Total']

for remov in remov_ls:
    ls_route.remove(remov)
# print('List of routes:', ls_route)

ls_route_str = [str(x) for x in ls_route]
print('List of routes:', ls_route_str)

# ls_route_str = ['5', '10', '15', '20', '30', '40', '50', '60', '70', '80', '100', 
#                 '110', '120', '130', '140', '150', '160', '240', '400', 
#                 '410', '420', '490', '500', '505', '515', '590', '595', '600'] 

# Define a function that create empty Sched Miles/Hours table in given dates ------------------------

def sched_table(change_date, end_date, ls_route_str):
    
    start = datetime.datetime.strptime(change_date, '%Y-%m-%d') # start date of a service change
    end = datetime.datetime.strptime(end_date, '%Y-%m-%d') # end date of a service change
    step = datetime.timedelta(days=1) # increase by 1 day at a time

    date_list = [] # create an empty list to record dates
    day_of_week_lst =[] # create an empty list to record day of the week
    service_type_list = [] # create an empty list to record service type (weekday, saturday or sunday)

    while start <= end: # when start date is early than (before) the end date

        day_of_week = datetime.datetime.strftime(start,'%A') # calculate the day of week from start date

        # calculate service type from day of week
        if day_of_week == 'Saturday': 
            service_type = 'Saturday'
            
        elif day_of_week == 'Sunday':
            service_type = 'Sunday'
            
        else:
            service_type = 'Weekday'

        date_list.append(str(start.date())) # append all days to the date list
        day_of_week_lst.append(day_of_week) # append all day of week to day of week list
        service_type_list.append(service_type) # append all service type of each day to service type list

        start += step # increase the start date by 1 day
        
    # compile the lists into a table, each list is a column
    df_sched_table = pd.DataFrame(list(zip(date_list, day_of_week_lst, service_type_list)),
               columns =['Date', 'Day of Week', 'Service Type']) # schedule table
    
    for route in ls_route_str: # add a new column to the table for every route 
        df_sched_table[route] = '' # no values
        
    return df_sched_table # return the empty schedule table


# Calculate Vehicle Revenue Miles and Hours -------------------------

# preparing Schedule Tables

df_sched_VRM = pd.DataFrame() # create an empty DataFrame object for recording one or more service change dataframes

# for each service IDs, create a sched miles and a sched hours table.

print('* Creating Scheduled Vehicle Revenue Miles / Hours templates...')

# loop through the 'service changes table' to check ID and dates
for index, row in df_service_change.iterrows():  # there are other faster methods to loop through the rows
    
    change_date = row['Change Date']  # extract the change date
    end_date = row['End Date']  # extract the end date
    change_id = row['Service Change ID'] # extract the service change ID
    
    # convert timestamp object to string 
    change_date = change_date.strftime("%Y-%m-%d")
    end_date = end_date.strftime("%Y-%m-%d")
    
    print('  -- start date / end date for Service Change', change_id, ':', change_date, '/', end_date)
    
    df_sched_miles = sched_table(change_date, end_date, ls_route_str) # create a df for each service change
    df_sched_miles['Service Change ID'] = change_id

    df_sched_VRM = df_sched_VRM.append(df_sched_miles, ignore_index = True) # combine tables from different service changes
    
df_sched_VRH = df_sched_VRM # Scheduled Vehicle Revenue Hours Table has the same structure as Scheduled Miles Table

# Update the miles and hours for Scheduled VRH and Scheduled VRM dataframe using the Excel schedule files
# 10_1_Scheduled Miles.xlsx and 10_2_Scheduled Hours.xlsx
# These weekly-schedule files are manually created by LeeTran service scheduler

print('* Updating Vehicle Revenue Miles templates...')
df_sched_VRM_copy = df_sched_VRM # make a copy for storing modified results
for index, row in df_sched_VRM.iterrows(): # loop through the template

    sched_VRM_date = row['Date'] # date is already in the template, extract the value
    sched_VRM_d_of_w = row['Day of Week'] # day of week is already in the template, extract the value
    sched_VRM_s_c_id = row['Service Change ID'] # service change ID is already in the template, extract the value
    
    for route in ls_route_str: # for every route
        
        # retrieve the scheduled miles of the route for a specific day from the manually created Excel schedule
        value = df_sched_mile_manual[int(route)].loc[(df_sched_mile_manual['Service Change ID'] == sched_VRM_s_c_id) & 
                                  (df_sched_mile_manual['Day of Week'] == sched_VRM_d_of_w)]  
    
        # replace the None value of the route for the specific day using mile value
        value_list = value.tolist() # convert the series to a list 
        value = value_list[0] # call the first element of the list which is the scheduled mile
        
        df_sched_VRM_copy.loc[(df_sched_VRM_copy.Date == sched_VRM_date), route] = value

print('* Exporting Scheduled Vehicle Revenue Miles Table 101_Sched_VRM.xlsx...')
df_sched_VRM_copy.to_excel("101_Sched_VRM.xlsx")

print('* Updating Vehicle Revenue Hours templates...')
df_sched_VRH_copy = df_sched_VRH # make a copy for storing modified results
for index, row in df_sched_VRH.iterrows():

    sched_VRH_date = row['Date'] 
    sched_VRH_d_of_w = row['Day of Week']
    sched_VRH_s_c_id = row['Service Change ID']
    
    for route in ls_route_str:
    # retrieve the hours of a cell of a route for a specific day from the scheduled hours Excel sheet
        value = df_sched_hour_manual[int(route)].loc[(df_sched_hour_manual['Service Change ID'] == sched_VRH_s_c_id) & 
                                  (df_sched_hour_manual['Day of Week'] == sched_VRH_d_of_w)]  
    
        # replace the None value of the route for the specific day using hour value
        value_list = value.tolist() # convert the series to a list 
        value = value_list[0] # call the first element of the list which is the scheduled hour
        
        df_sched_VRH_copy.loc[(df_sched_VRH_copy.Date == sched_VRH_date), route] = value
        
print('* Exporting Scheduled Vehicle Revenue Hours Table 102_Sched_VRH.xlsx ...')
df_sched_VRH_copy.to_excel("102_Sched_VRH.xlsx") 
print('* Done.')


# # Atypical Days
# 
# #### Atypical Day Definition by FTA (2022)
# A day on which the transit agency either: 
# - Does not operate its normal, regular schedule, or 
# - Provides extra service to meet demands for special events such as conventions, parades, or public celebrations, or operates significantly reduced service because of unusually bad weather (e.g., snowstorms, hurricanes, tornadoes, earthquakes) or major public disruptions (e.g., terrorism). 
# 
# Atypical days should not be included in the computation of average daily service.
# 
# #### Atypical days calculation in this workbook
# - Service miles and hours on an atypical day is first removed from the scheduled miles and hours
# - If not extra service on an atypical day, the actual miles and hous would be 0.
# - If service miles and hours occured on an atypical day, the service should be recorded only in 8_Added Runs table 
# - This program will add any service miles and hours on an atypical day based on the 8_Added Runs table in including deadheads
# - All the atypical days should not be used in calculating average daily service.

# In[2]:


# #### Processing Atypical Dates

print('* Removing atypical day miles and hours...')
atypical_list = df_atypical["Date"].tolist()

df_sched_VRM = pd.read_excel('101_Sched_VRM.xlsx', index_col=0)
df_sched_VRH = pd.read_excel('102_Sched_VRH.xlsx', index_col=0)

for a_date in atypical_list: # loop thourgh the atypical days
    a_date = a_date.strftime("%Y-%m-%d") # convert timestamp to string

    for route in ls_route_str: # loop thourgh the routes
        
        df_sched_VRM.loc[(df_sched_VRM.Date == a_date), route] = 0 # delete the miles/hours in the cells
        df_sched_VRH.loc[(df_sched_VRH.Date == a_date), route] = 0
        
        df_sched_VRM.loc[(df_sched_VRM.Date == a_date), 'Service Type'] = 'Atypical' # change the value in service type into 'Atypical'
        df_sched_VRH.loc[(df_sched_VRH.Date == a_date), 'Service Type'] = 'Atypical'

df_sched_VRM.to_excel("103_Sched_VRM_atypical_removed.xlsx") 
df_sched_VRH.to_excel("104_Sched_VRH_atypical_removed.xlsx") 
print('* Done.')


# # Added and Lost Runs
# - Record added miles and hours using positive numbers
# - Record lost miles and hours using negative numbers
# - Added and lost runs are originated from the daily logs of LeeTran Operation Department
# - There is another Python tool which automatically extracts these runs during a given date range from the logs
# - The tool is under S:\LeeTran\Planning\Technology\Python Projects\Retrieving Lost Trips from Daily Log

# In[3]:


# #### Processing Added and Lost Runs

print('* Reducing "Lost Runs"...')
df_VRM = pd.read_excel('103_Sched_VRM_atypical_removed.xlsx', index_col=0)
df_VRH = pd.read_excel('104_Sched_VRH_atypical_removed.xlsx', index_col=0)

# Loop through the lost records
for index, row in df_lost_run.iterrows():
    
    lost_date = row["Date"]
    lost_date = lost_date.strftime("%Y-%m-%d") # convert timestamp to string
    
    lost_route = row["Route"]
    lost_mile = float(row["Miles"])  # retrieve the lost miles
    lost_hour = float(row["Hours"])  # retrieve the lost hours
                     
    m = df_VRM.loc[df_VRM['Date'] == lost_date, str(lost_route)]
    m = m + lost_mile # subtracting the lost miles
    df_VRM.loc[(df_VRM.Date == lost_date), str(lost_route)] = m # update the mile value in the VRM table
    
    h = df_VRH.loc[df_VRH['Date'] == lost_date, str(lost_route)]
    h = h + lost_hour # subtracting the lost hours
    df_VRH.loc[(df_VRH.Date == lost_date), str(lost_route)] = h # update the mile value in the VRM table

df_VRM.to_excel("105_Sched_VRM_lost_removed.xlsx") 
df_VRH.to_excel("106_Sched_VRH_lost_removed.xlsx") 
print('* Done.')


print('* Adding "Added Runs"...')
df_sched_VRM = pd.read_excel('105_Sched_VRM_lost_removed.xlsx', index_col=0)
df_sched_VRH = pd.read_excel('106_Sched_VRH_lost_removed.xlsx', index_col=0)



# Loop through the added records
for index, row in df_added_run.iterrows():
    
    added_date = row["Date"]
    added_date = added_date.strftime("%Y-%m-%d") # convert timestamp to string

    added_route = row["Route"]
    added_mile = float(row["Miles"])  # retrieve the added miles
    added_hour = float(row["Hours"])  # retrieve the added hours
                      
    m = df_sched_VRM.loc[df_sched_VRM['Date'] == added_date, str(added_route)]
    m = m + added_mile # add the added miles
    df_sched_VRM.loc[(df_sched_VRM.Date == added_date), str(added_route)] = m # update the mile value in the VRM table
    
    h = df_sched_VRH.loc[df_sched_VRH['Date'] == added_date, str(added_route)]
    h = h + added_hour # add the added hours
    df_sched_VRH.loc[(df_sched_VRH.Date == added_date), str(added_route)] = h # update the mile value in the VRH table

df_sched_VRM.to_excel("10_Actual Vehicle Revenue Miles.xlsx") 
df_sched_VRH.to_excel("11_Actual Vehicle Revenue Hours.xlsx") 

print('* Done.')


# # Adding Deadhead Mile and Hour to Typical Days
# - This step processes the scheduled deadhead using the schedule Excel sheets
# - Deadhead of atypical days are considered in the next step

# In[4]:


### Total Vehicle Miles and Hours, processing deadhead

df_VRM = pd.read_excel('10_Actual Vehicle Revenue Miles.xlsx', index_col=0)
df_VRH = pd.read_excel('11_Actual Vehicle Revenue Hours.xlsx', index_col=0)

print("* Producing Actual Total Vehicle Miles and Hours tables...")

# Loop through the VRM records by day
for index, row in df_VRM.iterrows():
    
    day_of_week = row["Day of Week"]
    service_type = row["Service Type"]
    SC_id = row["Service Change ID"]
    date = row["Date"]

#     date = date_.strftime("%Y-%m-%d") # convert timestamp to string
    
    for route in ls_route_str: # loop thourgh the routes
        
        # search for the DHM of the day
        value_DHM = df_DHM[int(route)].loc[(df_DHM['Service Change ID'] == SC_id) & (df_DHM['Day of Week'] == day_of_week)]  
    
#         print('Route', route, 'DHM is:', float(value_DHM), 'on', service_type, day_of_week, 'in service ID', SC_id)

        # retrieve the actual revenue miles
        value_VRM = row[str(route)]
        updated_value_VRM = float(value_VRM) + float(value_DHM) # add the deadhed and get the actual hours 

        if str(service_type) != 'Atypical': # if the day is not an atypical day
            df_VRM.loc[(df_VRM.Date == date), str(route)] = float(updated_value_VRM) # update
        
df_VRM.to_excel("107_Total_Vehicle_Miles_deadhead_added.xlsx") 

# Loop through the VRH records by day
for index, row in df_VRH.iterrows():
    
    day_of_week = row["Day of Week"]
    service_type = row["Service Type"]
    SC_id = row["Service Change ID"]
    date = row["Date"]

    for route in ls_route_str: # loop thourgh the routes
        
        # search for the DHH of the day
        value_DHH = df_DHH[int(route)].loc[(df_DHH['Service Change ID'] == SC_id) & (df_DHH['Day of Week'] == day_of_week)]  
        # retrieve the actual revenue hours 
        value_VRH = row[str(route)]
        updated_value_VRH = float(value_VRH) + float(value_DHH) # add the deadhed and get the actual hours 

        if str(service_type) != 'Atypical': # if the day is not an atypical day
            df_VRH.loc[(df_VRH.Date == date), str(route)] = float(updated_value_VRH) # update
        
df_VRH.to_excel("108_Total_Vehicle_Hours_deadhead_added.xlsx") 
print('* Done.')


# # Processing Lost and Added Deadhead
# - This step processes the deadhead from Lost Runs and Added Runs
# - Deadhead of atypical days are considered here

# In[5]:


# The total vechile miles have not conisdered the lost and added deadhead

df_TVM = pd.read_excel('107_Total_Vehicle_Miles_deadhead_added.xlsx', index_col=0) 
df_TVH = pd.read_excel('108_Total_Vehicle_Hours_deadhead_added.xlsx', index_col=0)

print("* Processing deadhead miles and hours from 'Added Runs' and Lost Runs' ...")
        
# Loop through the added records to retrieve added deadhead miles and hours
for index, row in df_added_run.iterrows():
    
    added_date = row["Date"]
    added_date = added_date.strftime("%Y-%m-%d") # convert timestamp to string

    added_route = row["Route"]
    added_dh_mile = float(row["Deadhead Miles"])  # retrieve the added deadhead miles
    added_dh_hour = float(row["Deadhead Hours"])  # retrieve the added deadhead hours
                      
    m = df_TVM.loc[df_TVM['Date'] == added_date, str(added_route)] # check the original TVM miles of the day
    m = m + added_dh_mile # add the added deadhead miles
    df_TVM.loc[(df_TVM.Date == added_date), str(added_route)] = m # update the mile value in the TVM table
    
    h = df_TVH.loc[df_TVH['Date'] == added_date, str(added_route)]
    h = h + added_dh_hour # add the added deadhead hours
    df_TVH.loc[(df_TVH.Date == added_date), str(added_route)] = h # update the hour value in the TVM table
    
df_TVM.to_excel("109_Total_Vehicle_Miles_add_Added_Run_DH.xlsx") 
df_TVH.to_excel("110_Total_Vehicle_Hours_add_Added_Run_DH.xlsx") 

##--------------------------------------------------------------------------------------------------

df_TVM = pd.read_excel('109_Total_Vehicle_Miles_add_Added_Run_DH.xlsx', index_col=0)
df_TVH = pd.read_excel('110_Total_Vehicle_Hours_add_Added_Run_DH.xlsx', index_col=0)

# Loop through the added records to retrieve added deadhead miles and hours
for index, row in df_lost_run.iterrows():
    
    lost_date = row["Date"] # retrieve the date of the lost record
    lost_date = lost_date.strftime("%Y-%m-%d") # convert timestamp to string

    lost_route = row["Route"]
    lost_dh_mile = float(row["Deadhead Miles"])  # retrieve the lost deadhead miles
    lost_dh_hour = float(row["Deadhead Hours"])  # retrieve the lost deadhead hours
                      
    m = df_TVM.loc[df_TVM['Date'] == lost_date, str(lost_route)] # check the original TVM miles of the day
    m = m + lost_dh_mile # add the lost deadhead miles
    df_TVM.loc[(df_TVM.Date == lost_date), str(lost_route)] = m # update the mile value in the TVM table
    
    h = df_TVH.loc[df_TVH['Date'] == lost_date, str(lost_route)]
    h = h + lost_dh_hour # add the lost deadhead hours
    df_TVH.loc[(df_TVH.Date == lost_date), str(lost_route)] = h # update the hour value in the TVM table
    
df_TVM.to_excel("12_Actual Total Vehicle Miles.xlsx") 
df_TVH.to_excel("13_Actual Total Vehicle Hours.xlsx") 

print('* Done')
# input("Press Enter to exit...")


# # MR-20 Calculation
# 
# - The output is for FTA NTD monthly fixed-route risdership report 
# - Update the information in Excel table 1_Daily Ridership by Route.xlsx and run the script
# - MR-20 includes the data from atypical days

# In[59]:


print()
print('*****************************************************')
print('-----------MR-20-------------')
# Read ridership data
df_ridership = pd.read_excel('1_Daily Ridership by Route.xlsx') # Daily Ridership by Route
df_ridership['year_month'] = pd.to_datetime(df_ridership['Date']).dt.to_period('M')

# Sum by month
month_rider = df_ridership.groupby("year_month").sum()
month_rider = month_rider.rename(columns={"Total": "UPT"})

# Read VRM data
df = pd.read_excel('10_Actual Vehicle Revenue Miles.xlsx', index_col=0)
df['year_month'] = pd.to_datetime(df['Date']).dt.to_period('M')
df = df.drop(['Service Change ID'], axis=1)
month_vrm = df.groupby("year_month").sum()
month_vrm['total_vrm'] = month_vrm.sum(axis=1)

# # Read VRH data
df = pd.read_excel('11_Actual Vehicle Revenue Hours.xlsx', index_col=0)
df['year_month'] = pd.to_datetime(df['Date']).dt.to_period('M')
df = df.drop(['Service Change ID'], axis=1)
month_vrh = df.groupby("year_month").sum()
month_vrh['total_vrh'] = month_vrh.sum(axis=1)

MR20 = pd.merge(month_rider, month_vrm[["total_vrm"]], on="year_month", how="left")
MR20 = pd.merge(MR20, month_vrh[["total_vrh"]], on="year_month", how="left")
MR20 = MR20.rename(columns={"Total": "UPT", "total_vrm": "VRM", "total_vrh": "VRH"})
MR20 = MR20[["UPT", "VRM", "VRH"]]
MR20 = MR20.reset_index() # reset index


df = pd.read_excel('0_VOMs.xlsx')
for col in df.columns:   # change all the column names to string
    df.rename(columns = {col:str(col)}, inplace = True)

for route_num in ls_route_str:   # change all the values under buses to number
#     print(route_num)
    df[route_num] = df[route_num].astype(int)
    
#specify the columns to sum
# print(ls_route_str)

#find sum of columns specified 
df['daily_VOM'] = df[ls_route_str].sum(axis=1)

df['year_month'] = pd.to_datetime(df['Date']).dt.to_period('M') # add a new column - month of year

monthly_df = df.groupby("year_month")  # Group by "year_month" column
monthly_df = monthly_df.max() # Get maximum values in each group
monthly_df = monthly_df.reset_index() # reset index
monthly_df = monthly_df[['year_month','daily_VOM']]

MR20_joined = pd.merge(MR20, monthly_df, on='year_month', how='left')
MR20_joined = MR20_joined.rename(columns={"daily_VOM": "VOM"})
MR20_joined = MR20_joined.round({"VRM":2, "VRH":2})

MR20_joined.to_excel("MR-20.xlsx", index = False) 

print(MR20_joined)
print()
print('MR-20.xlsx is saved in the folder.')
print('*****************************************************')


# # S-10 Calculation

# ### - S10 Service Supplied

# In[7]:


print()
print('********************  S-10   **********************')
print()
print('*****************************************************')


# calculating average Weekday VOMS, average Saturday VOMS, average Sunday VOMS
df = pd.read_excel('0_VOMs.xlsx')
for col in df.columns:   # change all the column names to string
    df.rename(columns = {col:str(col)}, inplace = True)

for route_num in ls_route_str:   # change all the values under buses to number
#     print(route_num)
    df[route_num] = df[route_num].astype(int)
    
#specify the columns to sum
# print(ls_route_str)
#find sum of columns specified 
df['daily_VOM'] = df[ls_route_str].sum(axis=1)

# Extract weekday
# weekday_df.dtypes  # check each column's data type
weekday_df = df.loc[df['Service Type'] == 'Weekday']
weekday_df.set_index('Date', inplace = True)
wd_sc_id1 = weekday_df['2021-10-01' : '2021-11-20']
wd_sc_id1_VOMS = wd_sc_id1['daily_VOM'].max()
wd_sc_id1_num = len(wd_sc_id1)

wd_sc_id2 = weekday_df['2021-11-21' : '2022-1-1']
wd_sc_id2_VOMS = wd_sc_id2['daily_VOM'].max()
wd_sc_id2_num = len(wd_sc_id2)

wd_sc_id3 = weekday_df['2022-1-2' : '2022-2-25']
wd_sc_id3_VOMS = wd_sc_id3['daily_VOM'].max()
wd_sc_id3_num = len(wd_sc_id3)

wd_sc_id4 = weekday_df['2022-2-26' : '2022-4-23']
wd_sc_id4_VOMS = wd_sc_id4['daily_VOM'].max()
wd_sc_id4_num = len(wd_sc_id4)

wd_sc_id5 = weekday_df['2022-4-24' : '2022-9-30']
wd_sc_id5_VOMS = wd_sc_id5['daily_VOM'].max()
wd_sc_id5_num = len(wd_sc_id5)

# Extract Saturday
sat_df = df.loc[df['Service Type'] == 'Saturday']
sat_df.set_index('Date', inplace = True)
sat_sc_id1 = sat_df['2021-10-01' : '2021-11-20']
sat_sc_id1_VOMS = sat_sc_id1['daily_VOM'].max()
sat_sc_id1_num = len(sat_sc_id1)

sat_sc_id2 = sat_df['2021-11-21' : '2022-1-1']
sat_sc_id2_VOMS = sat_sc_id2['daily_VOM'].max()
sat_sc_id2_num = len(sat_sc_id2)

sat_sc_id3 = sat_df['2022-1-2' : '2022-2-25']
sat_sc_id3_VOMS = sat_sc_id3['daily_VOM'].max()
sat_sc_id3_num = len(sat_sc_id3)

sat_sc_id4 = sat_df['2022-2-26' : '2022-4-23']
sat_sc_id4_VOMS = sat_sc_id4['daily_VOM'].max()
sat_sc_id4_num = len(sat_sc_id4)

sat_sc_id5 = sat_df['2022-4-24' : '2022-9-30']
sat_sc_id5_VOMS = sat_sc_id5['daily_VOM'].max()
sat_sc_id5_num = len(sat_sc_id5)

# Extract Sunday
san_df = df.loc[df['Service Type'] == 'Sunday']
san_df.set_index('Date', inplace = True)
san_sc_id1 = san_df['2021-10-01' : '2021-11-20']
san_sc_id1_VOMS = san_sc_id1['daily_VOM'].max()
san_sc_id1_num = len(san_sc_id1)

san_sc_id2 = san_df['2021-11-21' : '2022-1-1']
san_sc_id2_VOMS = san_sc_id2['daily_VOM'].max()
san_sc_id2_num = len(san_sc_id2)

san_sc_id3 = san_df['2022-1-2' : '2022-2-25']
san_sc_id3_VOMS = san_sc_id3['daily_VOM'].max()
san_sc_id3_num = len(san_sc_id3)

san_sc_id4 = san_df['2022-2-26' : '2022-4-23']
san_sc_id4_VOMS = san_sc_id4['daily_VOM'].max()
san_sc_id4_num = len(san_sc_id4)

san_sc_id5 = san_df['2022-4-24' : '2022-9-30']
san_sc_id5_VOMS = san_sc_id5['daily_VOM'].max()
san_sc_id5_num = len(san_sc_id5)

# Extract Atypical
atypical_df = df.loc[df['Service Type'] == 'Atypical']
atypical_df.set_index('Date', inplace = True)
atypical_sc_id1 = atypical_df['2021-10-01' : '2021-11-20']
atypical_sc_id1_num = len(atypical_sc_id1)

atypical_sc_id2 = atypical_df['2021-11-21' : '2022-1-1']
atypical_sc_id2_num = len(atypical_sc_id2)

atypical_sc_id3 = atypical_df['2022-1-2' : '2022-2-25']
atypical_sc_id3_num = len(atypical_sc_id3)

atypical_sc_id4 = atypical_df['2022-2-26' : '2022-4-23']
atypical_sc_id4_num = len(atypical_sc_id4)

atypical_sc_id5 = atypical_df['2022-4-24' : '2022-9-30']
atypical_sc_id5_num = len(atypical_sc_id5)

# Preparing service changes table
df_sc = pd.DataFrame() # create an empty DataFrame object

# Assign data of lists
sc_data = {'Service Change ID':[1,2,3,4,5], 
           'Change Date': ['10/1/2021', '11/21/2021', '1/2/2022', '2/26/2022', '4/24/2022'],
           'End Date':   ['11/20/2021', '1/1/2022', '2/25/2022', '4/23/2022', '9/30/2022'],
           'Weekdays': [wd_sc_id1_num, wd_sc_id2_num, wd_sc_id3_num, wd_sc_id4_num, wd_sc_id5_num],
           'Sat': [sat_sc_id1_num, sat_sc_id2_num, sat_sc_id3_num, sat_sc_id4_num, sat_sc_id5_num],
           'Sun': [san_sc_id1_num, san_sc_id2_num, san_sc_id3_num, san_sc_id4_num, san_sc_id5_num],
           'Atypical': [atypical_sc_id1_num, atypical_sc_id2_num, atypical_sc_id3_num, atypical_sc_id4_num, atypical_sc_id5_num],
           'WD VOMS': [wd_sc_id1_VOMS, wd_sc_id2_VOMS, wd_sc_id3_VOMS, wd_sc_id4_VOMS, wd_sc_id5_VOMS],
           'SA VOMS': [sat_sc_id1_VOMS, sat_sc_id2_VOMS, sat_sc_id3_VOMS, sat_sc_id4_VOMS, sat_sc_id5_VOMS],
           'SU VOMS': [san_sc_id1_VOMS, san_sc_id2_VOMS, san_sc_id3_VOMS, san_sc_id4_VOMS, san_sc_id5_VOMS]
           }

df_sc = pd.DataFrame(sc_data)
df_sc.to_excel('14_Service Changes VOMS.xlsx', index = False)

avg_WD_VOMS = round(np.sum(df_sc['Weekdays'] * df_sc['WD VOMS'])/(np.sum(df_sc['Weekdays'])), 2)
avg_SA_VOMS = round(np.sum(df_sc['Sat'] * df_sc['SA VOMS'])/(np.sum(df_sc['Sat'])), 2)
avg_SU_VOMS = round(np.sum(df_sc['Sun'] * df_sc['SU VOMS'])/(np.sum(df_sc['Sun'])), 2)

df_VOMS_data = {'Service Type': ['Atypical', 'Saturday', 'Sunday', 'Weekday', 'Annual'],
                'Average VOMS': ['N/A', avg_SA_VOMS, avg_SU_VOMS, avg_WD_VOMS, 'N/A']}
df_VOMS = pd.DataFrame(df_VOMS_data)

# Total Actual Vehicle Miles
df = pd.read_excel('12_Actual Total Vehicle Miles.xlsx', index_col=0)
# df.drop('Service Change ID', inplace=True, axis=1)

sum_column = 0
for route in ls_route_str: # Sum up daily
    sum_column = sum_column + df[route] 

df['SUM_DAILY_df'] = sum_column
print(' ----  Average AVM by service type ----')
AVM_average = df[["Service Type", "SUM_DAILY_df"]].groupby("Service Type").mean().round(2) # The average of each service type
print(AVM_average)
print()

AVM_average.reset_index(inplace = True)
AVM_average = AVM_average.rename(columns = {'index': "Service Type"})
# list(AVM_average.columns)

AVM_Annual = df['SUM_DAILY_df'].sum()
print('Annual AVM:', round(AVM_Annual, 2))

df2 = ['Annual', round(AVM_Annual, 2)]
df2_series = pd.Series(df2, index = AVM_average.columns)

AVM_average = AVM_average.append(df2_series, ignore_index = True)
AVM_average = AVM_average.rename(columns = {'SUM_DAILY_df': "Average AVM"})

print()
print('*****************************************************')
print()

# Total Actual Vehicle Revenue Miles
df = pd.read_excel('10_Actual Vehicle Revenue Miles.xlsx', index_col=0)

sum_column = 0
for route in ls_route_str: # Sum up daily
    sum_column = sum_column + df[route] 

df['SUM_DAILY_df'] = sum_column
print(' ----  Average AVRM by service type ----')
AVRM_average = df[["Service Type", "SUM_DAILY_df"]].groupby("Service Type").mean().round(2) # The average of each service type
print(AVRM_average)

print()
AVRM_Annual = df['SUM_DAILY_df'].sum()
print('Annual AVRM:', round(AVRM_Annual, 2))

AVRM_average.reset_index(inplace = True)
AVRM_average = AVRM_average.rename(columns = {'index': "Service Type"})
df3 = ['Annual', round(AVRM_Annual, 2)]
df3_series = pd.Series(df3, index = AVRM_average.columns)
AVRM_average = AVRM_average.append(df3_series, ignore_index = True)
AVRM_average = AVRM_average.rename(columns = {'SUM_DAILY_df': "Average AVRM"})

merged1= AVM_average.merge(AVRM_average, on="Service Type", how="left")

print()
print('*****************************************************')
print()


# In[8]:


# Total Actual Vehicle Hours
df = pd.read_excel('13_Actual Total Vehicle Hours.xlsx', index_col=0)

sum_column = 0
for route in ls_route_str: # Sum up daily
    sum_column = sum_column + df[route] 

df['SUM_DAILY_df'] = sum_column
print(' ----  Average AVH by service type ----')
AVH_average = df[["Service Type", "SUM_DAILY_df"]].groupby("Service Type").mean().round(2) # The average of each service type
print(AVH_average)

print()
AVH_Annual = df['SUM_DAILY_df'].sum()
print('Annual AVH:', round(AVH_Annual,2))

AVH_average.reset_index(inplace = True)
AVH_average = AVH_average.rename(columns = {'index': "Service Type"})
df4 = ['Annual', round(AVH_Annual, 2)]
df4_series = pd.Series(df4, index = AVH_average.columns)
AVH_average = AVH_average.append(df4_series, ignore_index = True)
AVH_average = AVH_average.rename(columns = {'SUM_DAILY_df': "Average AVH"})

merged2= merged1.merge(AVH_average, on="Service Type", how="left")

print()
print('*****************************************************')
print()


# Total Actual Vehicle Revenue Miles
df = pd.read_excel('11_Actual Vehicle Revenue Hours.xlsx', index_col=0)

sum_column = 0
for route in ls_route_str: # Sum up daily
    sum_column = sum_column + df[route] 

df['SUM_DAILY_df'] = sum_column
print(' ----  Average AVRH by service type ----')
AVRH_average = df[["Service Type", "SUM_DAILY_df"]].groupby("Service Type").mean().round(2) # The average of each service type
print(AVRH_average)

print()
AVRH_Annual = df['SUM_DAILY_df'].sum()
print('Annual AVRH:', round(AVRH_Annual,2))

AVRH_average.reset_index(inplace = True)
AVRH_average = AVRH_average.rename(columns = {'index': "Service Type"})
df5 = ['Annual', round(AVRH_Annual, 2)]
df5_series = pd.Series(df5, index = AVRH_average.columns)
AVRH_average = AVRH_average.append(df5_series, ignore_index = True)
AVRH_average = AVRH_average.rename(columns = {'SUM_DAILY_df': "Average AVRH"})

merged3= merged2.merge(AVRH_average, on="Service Type", how="left")

print()
print('*****************************************************')
print()


# print('Deadhead Miles (Average Weekday):', )
# print(' ---- Deadhead ---- ')
# print('Deadhead Miles (Annual Total):', round(AVM_Annual - AVRM_Annual, 2))


# print('Deadhead Hours (Average Weekday):', )
# print('Deadhead Hours (Average Saturday):', )
# print('Deadhead Hours (Average Sunday):', )
# print('Deadhead Hours (Annual Total):', round(AVH_Annual - AVRH_Annual, 2))


# ### - S10 DeadHead
# - Deadhead miles/hours = Actual vehicle miles/hours - Acutal vehicle revenue miles/hours

# In[9]:


print(' ---- Deadhead ---- ')
print()
print('Deadhead Miles (Annual Total):', round(AVM_Annual - AVRM_Annual, 2))

# print('Deadhead Hours (Average Weekday):', )
# print('Deadhead Hours (Average Saturday):', )
# print('Deadhead Hours (Average Sunday):', )
print('Deadhead Hours (Annual Total):', round(AVH_Annual - AVRH_Annual, 2))
print()
print('*****************************************************')
print()


# In[10]:


# Total Scheduled Vehicle Revenue Miles in S10

df = pd.read_excel('103_Sched_VRM_atypical_removed.xlsx', index_col=0)
sum_column = 0
for route in ls_route_str: # Sum up daily
    sum_column = sum_column + df[route] 

df['SUM_DAILY_df'] = sum_column
print(' ----  Average Sched VRM by service type ----')
Sched_VRM_average = df[["Service Type", "SUM_DAILY_df"]].groupby("Service Type").mean().round(2) # The average of each service type
print(Sched_VRM_average)

print()
Sched_VRM_Annual = df['SUM_DAILY_df'].sum()
print('Annual Scheduled VRM:', round(Sched_VRM_Annual,2))

Sched_VRM_average.reset_index(inplace = True)
Sched_VRM_average = Sched_VRM_average.rename(columns = {'index': "Service Type"})
df6 = ['Annual', round(Sched_VRM_Annual, 2)]
df6_series = pd.Series(df6, index = Sched_VRM_average.columns)
Sched_VRM_average = Sched_VRM_average.append(df6_series, ignore_index = True)
Sched_VRM_average = Sched_VRM_average.rename(columns = {'SUM_DAILY_df': "Average Sched VRM"})

merged4= merged3.merge(Sched_VRM_average, on="Service Type", how="left")

print()
print('*****************************************************')
print()


# In[11]:


# Unlinked Passenger Trips (UPT)
print('---- Average Unlinked Passenger Trips (UPT) ----')
# df = pd.read_excel('1_Daily Ridership by Route.xlsx', index_col=0)
df = pd.read_excel('1_Daily Ridership by Route.xlsx')

# sum_column = 0
# for route in ls_route_str: # Sum up daily
#     sum_column = sum_column + df[route] 

# df['SUM_DAILY_df'] = sum_column

# print(' ----  Average Sched VRM by service type ----')
UPT = df[["Service Type", "Total"]].groupby("Service Type").mean().round(2) # The average of each service type
print(UPT)

print()
UPT_Annual = df['Total'].sum()
print('Annual Total UPT:', UPT_Annual)

UPT.reset_index(inplace = True)
UPT = UPT.rename(columns = {'index': "Service Type"})
df7 = ['Annual', UPT_Annual]
df7_series = pd.Series(df7, index = UPT.columns)
UPT = UPT.append(df7_series, ignore_index = True)
UPT = UPT.rename(columns = {'Total': "Average UPT"})

merged5= merged4.merge(UPT, on="Service Type", how="left")

print()
print('*****************************************************')
print()

print('---- Total Unlinked Passenger Trips (UPT) ----')
UPT_sum = df[["Service Type", "Total"]].groupby("Service Type").sum() # The average of each service type
print(UPT_sum)
print()
print('Annual Total UPT:', UPT_Annual)

UPT_sum.reset_index(inplace = True)
UPT_sum = UPT_sum.rename(columns = {'index': "Service Type"})
df8 = ['Annual', UPT_Annual]
df8_series = pd.Series(df8, index = UPT_sum.columns)
UPT_sum = UPT_sum.append(df8_series, ignore_index = True)
UPT_sum = UPT_sum.rename(columns = {'Total': "Total UPT"})

merged6= merged5.merge(UPT_sum, on="Service Type", how="left")

print()
print('*****************************************************')
print()


# In[12]:


# Days Operated
print('---- Days Operated ----')
df = pd.read_excel('1_Daily Ridership by Route.xlsx')

UPT_count = df[["Service Type", "Total"]].groupby("Service Type").count() # The average of each service type
print(UPT_count)

UPT_count.reset_index(inplace = True)
UPT_count = UPT_count.rename(columns = {'index': "Service Type"})
df9 = ['Annual', UPT_count['Total'].sum()]
df9_series = pd.Series(df9, index = UPT_count.columns)
UPT_count = UPT_count.append(df9_series, ignore_index = True)
UPT_count = UPT_count.rename(columns = {'Total': "Days Operated"})

merged7 = merged6.merge(UPT_count, on="Service Type", how="left")

print()
index = df.index
number_of_rows = len(index)
print('Number of rows in Daily Ridership Sheet:', number_of_rows)
print()


df10 = ['N/A', 'N/A', 'N/A', 'N/A', round(AVM_Annual - AVRM_Annual, 2)]
merged7['Deadhead Miles (Annual Total)'] = df10

df11 = ['N/A', 'N/A', 'N/A', 'N/A', round(AVH_Annual - AVRH_Annual, 2)]
merged7['Deadhead Hours (Annual Total)'] = df11

df12 = ['N/A', 'N/A', 'N/A', 'N/A', number_of_rows]
merged7['Number of rows in Daily Ridership Sheet'] = df12

# Merged with average VOMS
merged8 = merged7.merge(df_VOMS, on="Service Type", how="left")

# Create final S-10
S_10 = merged8.T
S_10 = S_10.rename(index = {'Average AVM': 'Total Actual Vehicle Miles', 
                            'Average AVRM': 'Total Actual Vehicle Revenue Miles',
                            'Average AVH': 'Total Actual Vehicle Hours',
                            'Average AVRH': 'Total Actual Vehicle Revenue Hours',
                            'Average Sched VRM': 'Total Scheduled Vehicle Revenue Miles',
                            'Average UPT': 'Average Unlinked Passenger Trips (UPT)',
                            'Total UPT': 'Total Unlinked Passenger Trips (UPT)',
                            'Deadhead Miles (Annual Total)': 'Deadhead Miles',
                            'Deadhead Hours (Annual Total)': 'Deadhead Hours',
                            'Average VOMS': 'Vehicles in Operation (VOMS)'
                            })

# S_10 = S_10.iloc[1:, :]
S_10 = S_10.rename(columns = {'Atypical': 'Average Atypical Schedule',
                              'Saturday': 'Average Saturday Schedule',
                              'Sunday': 'Average Sunday Schedule',
                              'Weekday': 'Average Weekday Schedule',
                              'Annual': 'Annual Total'
                              })

S_10 = S_10.rename(columns = {0: 'Average Atypical Schedule',
                              1: 'Average Saturday Schedule',
                              2: 'Average Sunday Schedule',
                              3: 'Average Weekday Schedule',
                              4: 'Annual Total'
                              })

S_10 = S_10.reindex(columns = ['Average Weekday Schedule', 'Average Saturday Schedule', 
                               'Average Sunday Schedule', 'Average Atypical Schedule', 'Annual Total'])

S_10 = S_10.style.set_properties( **{'text-align': 'right'})

# Save the final dataframe into excel
S_10.to_excel('S-10.xlsx') 

print()
print('S-10.xlsx is saved in the folder.')
print('*****************************************************')
print()


input("Press Enter to continue...")

