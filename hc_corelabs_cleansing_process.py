# -*- coding: utf-8 -*-
"""
Created on Sun Apr 30 09:33:41 2023

@author: londoncm
"""

import pandas as pd
import os

# Adding the path where is the data located

os.chdir(r'Z:\Python\Test\HC_data') 

os.listdir()

# --------------Import HC Data (Excel)-------------------------

df_cl1 = pd.read_excel('hc_corelabs_march_2023.xlsx', dtype=object) # Change file name every month

df_cl1.head()
df_cl1.dtypes

# Remove rows when department is Research Center & Reserach Office
df_cl1 = df_cl1.loc[df_cl1['RH3 Department'] == 'Core Labs', :]





# Null reeplacement and new columns

df_cl1['RH3 Division'] = df_cl1['RH3 Division'].replace({"Vice President Resea": 3})
df_cl1['RH3 Department'] = df_cl1['RH3 Department'].replace({"Core Labs": 1020})
df_cl1['Cost Center'] = df_cl1['Cost Center'].replace({"#": 0})
df_cl1['Supervisor'] = df_cl1['Supervisor'].replace({"#": '100000'})
df_cl1['Classification'] = df_cl1['Classification'].replace({"PTSA": "DSFTEC", 
                                                       "Post-Doctoral":"FTE/Postdoc",
                                                       "Research":"Research/Engineer",
                                                       "Staff Scientist-Engineer":"Research/Engineer"})

df_cl1['Grade'] = df_cl1['Grade'].replace({"1R": "R1", "2R":"R2", "3R":"R3", "4R":"R4",
                                                       "5R":"R5", "6R":"R6",
                                                       "1P":"P1", "2P":"P2",
                                                       "3P":"P3", "4P":"P4",
                                                       "5P":"P5", "6P":"P6",
                                                       "7P":"P7", "8P":"P8"})


def cost_center(row):
    if row['Cost Center'] == 0:
       row['Cost Center'] = row['Master Cost Center']
    else:
        row['Cost Center'] = row['Cost Center']
    return row 


df_cl1 = df_cl1.apply(cost_center, axis=1)

def cc_desc(row):
    if row['CC Desc'] == '1000/Not assigned':
       row['CC Desc'] = row['MCC Desc']
    else:
        row['CC Desc'] = row['CC Desc']
    return row 


df_cl1 = df_cl1.apply(cc_desc, axis=1)



def Type_Position(row):
    
    if row['Classification']== 'Professional' or row['Classification']== 'Research/Engineer'\
        or row['Classification']== 'Technician' or row['Classification']== 'Manager'\
        or row['Classification']== 'Support':    
        row['Type Position'] = 'Onboard Position'
    elif row['Classification']== 'DSFTEC' or row['Classification']== 'Vendor Manpower':
        row['Type Position'] = 'Other Labor'
    elif row['Classification']== 'Faculty' or row['Classification']== 'Independent Consultant'\
        or row['Classification']== 'Saudi Development Program':
        row['Type Position'] = 'Other Miscellaneous'
    elif row['Classification']== 'FTE/Postdoc':
        row['Type Position'] = 'Wrongly Charged'
    elif row['Classification']== 'Executive':
        row['Type Position'] = 'Wrongly Charged'
    else:
       row['Type Position'] = ''
    return row   


df_cl1 = df_cl1.apply(Type_Position, axis=1)

df_cl1['CC Desc'] = df_cl1['CC Desc'].replace({"CL - Ops and Support": 'Operation & Support',\
                                               "CL - NanoFab": 'Nanofabrication',
                                               'CL - KVL': 'Visualization',
                                               'CL - BCL': 'Bioscience',
                                               'CL - CMOR': 'Coastal & Marine Resources',
                                               'CL - KSL': 'Supercomputing',
                                               'CL - ACL':'Analytical Chemistry',
                                               'CL - CW': 'Central Workshop',
                                               'CL - I&C': 'Imaging & Characterization',
                                               'CL - LEM': 'LEM',
                                               'CL - GH': 'Greenhouse',
                                               'CL - ARF': 'Animal Resources Facility',
                                               'CL - RLF': 'Radiation Labelling Facility'
                                               })

# Remove colummns we do not use in hc Data

df_cl1 = df_cl1.drop([
    'WBS Element',
    'RH2 Department',
    'WBS Desc',
    'Employee Group',
    'Organizational Unit',
    'OU Desc',
    'Master Cost Center',
    'Job',
    'MCC Desc',
    'User Name',
    'Contract End Date',
    ], axis=1)

# Change type of column

df_cl1['Actual'] = df_cl1['Actual'].astype('float64')
df_cl1['Cost Center'] = df_cl1['Cost Center'].astype('int64')

df_cl1.dtypes



# Table with Onboard positions

df_cl2 = df_cl1.loc[df_cl1['Type Position'] == 'Onboard Position', :]

df_cl2.dtypes

df_cl3 = df_cl1.loc[df_cl1['Type Position'] != 'Onboard Position', :]


# --------------Import Plan HC Data 2022-23 (Excel)-------------------------

plan_data_cl = pd.read_excel('Plan_CoreLabs_2023.xlsx', sheet_name='plan_cl',
                    dtype=object) 


plan_data_cl.head()
plan_data_cl.dtypes

plan_data_cl['RH3 Division'] = plan_data_cl['RH3 Division'].astype('int64')
plan_data_cl['RH3 Department'] = plan_data_cl['RH3 Department'].astype('int64')
plan_data_cl['Plan'] = plan_data_cl['Plan'].astype('float64')
plan_data_cl['Position'] = plan_data_cl['Position'].astype('str')
plan_data_cl['Cost Center'] = plan_data_cl['Cost Center'].astype('int64')


# plan_data_cl['Actual_Plan'] = "Plan"


groupby_actual = df_cl2.groupby(by=["Cost Center", "Classification"]).sum("Actual")

# groupby_actual = groupby_actual.drop([
#     'RH3 Division',
#     'RH3 Department',
#     ], axis=1)


groupby_plan = plan_data_cl.groupby(by=["Cost Center", "Classification"]).sum("Actual")

# groupby_plan = groupby_plan.drop([
#     'RH3 Division',
#     'RH3 Department',
#     ], axis=1)


groupby_plan['Actual'] = groupby_plan['Plan'] - groupby_actual['Actual']

groupby_plan['Actual'].fillna(1, inplace=True)

groupby_plan = groupby_plan.loc[groupby_plan['Actual'] != 0, :]

groupby_plan = groupby_plan.drop([
    'Plan',
    ], axis=1)


groupby_plan['RH3 Division'] = 3
groupby_plan['RH3 Department'] = 1020


vacants = groupby_plan.reset_index()


def cc_des(row):
    if row['Cost Center'] == 12100 :    
        row['CC Desc'] =  'Operation & Support'
    elif row['Cost Center'] == 12110 :    
        row['CC Desc'] =  'Nanofabrication'
    elif row['Cost Center'] == 12120 :    
        row['CC Desc'] =  'Visualization'
    elif row['Cost Center'] == 12130 :    
        row['CC Desc'] =  'Bioscience'
    elif row['Cost Center'] == 12140 :    
        row['CC Desc'] =  'Coastal & Marine Resources'
    elif row['Cost Center'] == 12150 :    
        row['CC Desc'] =  'Supercomputing'
    elif row['Cost Center'] == 12160 :    
        row['CC Desc'] =  'Analytical Chemistry'
    elif row['Cost Center'] == 12170 :    
        row['CC Desc'] =  'Central Workshop'
    elif row['Cost Center'] == 12190 :    
        row['CC Desc'] =  'Imaging & Characterization'
    elif row['Cost Center'] == 12360 :    
        row['CC Desc'] =  'LEM'
    elif row['Cost Center'] == 12380 :    
        row['CC Desc'] =  'Greenhouse'
    elif row['Cost Center'] == 12390 :    
        row['CC Desc'] =  'Animal Resources Facility'
    elif row['Cost Center'] == 12400 :    
        row['CC Desc'] =  'Radiation Labelling Facility'
    else:
       row['Cost Center'] =  'NA'
    return row 

vacants = vacants.apply(cc_des, axis=1)

vacants['Grade'] = '0'
vacants['KAUSTID'] = '100000'
vacants['Employee Name'] = 'Vacant Position'
vacants['Position'] = '0'
vacants['Job Title'] = vacants['Classification']
vacants['Supervisor'] = '100000'
vacants['Entry Date'] = 'Not assigned'
vacants['Type Position'] = 'Vacant Position'

# Change the columns order

vacants = vacants[[
    'RH3 Division',
    'RH3 Department',
    'Cost Center',
    'CC Desc',
    'Classification',
    'Grade',
    'KAUSTID',
    'Employee Name',
    'Position',
    'Job Title',
    'Supervisor',
    'Entry Date',
    'Actual',
    'Type Position',
    ]]

# Join Onboard Table with vacant table

df_cl4 = pd.concat([df_cl2, vacants])

hc_data_cl = pd.concat([df_cl4, df_cl3])

# Create new columns

hc_data_cl['Aprroved Position'] = '0'
hc_data_cl['Role'] = 'Not assigned'
hc_data_cl['Approved Grade'] = '0'


def actual_position2(row):
    
    if row['Type Position']== 'Vacant Position':    
        row['Actual Position'] = '0'
    else:
       row['Actual Position'] = row['Position']
    return row   

hc_data_cl = hc_data_cl.apply(actual_position2, axis=1)


def approve_classification2(row):
    
    if row['Type Position']== 'Onboard Position' or row['Type Position']== 'Vacant Position':    
        row['Approved Classification'] = row['Classification']
    else:
       row['Approved Classification'] = '0'
    return row   

hc_data_cl = hc_data_cl.apply(approve_classification2, axis=1)


def position2(row):
    
    if row['Type Position']== 'Onboard Position' or row['Type Position']== 'Vacant Position':    
        row['Position#'] = 1
    elif row['Type Position']== 'Other Labor':
        row['Position#'] = 2
    elif row['Type Position']== 'Wrongly Charged':
        row['Position#'] = 3   
    else:
       row['Position#'] = 4
    return row 

hc_data_cl = hc_data_cl.apply(position2, axis=1)



def esg_reclassified2(row):
    
    if row['Type Position']==  'Vacant Position':    
        row['ESGReclassified_ID'] =  '0'
    else:
       row['ESGReclassified_ID'] = row['Grade'] 
    return row 


hc_data_cl = hc_data_cl.apply(esg_reclassified2, axis=1)

def plan2(row):
    
    if row['Type Position']== 'Onboard Position' or row['Type Position']== 'Vacant Position':     
        row['Plan'] = row['Actual'] 
    else:
       row['Plan'] = 0
    return row 

hc_data_cl = hc_data_cl.apply(plan2, axis=1)



# Remove columns

hc_data_cl = hc_data_cl.drop([
    'Grade',
    'Position',
    ], axis=1)

# Rename Columns

hc_data_cl = hc_data_cl.rename(columns={
    'RH3 Division': 'Division_ID',
    'RH3 Department': 'Department_ID',
    'Cost Center': 'CostCenter_ID',
    'Position#':'Position',
    'Supervisor':'Supervisor_ID'
    })


# Change the columns order

hc_data_cl = hc_data_cl[[
    'Division_ID',
    'Department_ID',
    'CostCenter_ID',
    'KAUSTID',
    'Employee Name',
    'Supervisor_ID',
    'Actual Position',
    'Aprroved Position',
    'Classification',
    'Approved Classification',
    'Approved Grade',
    'ESGReclassified_ID',
    'Type Position',
    'Position',
    'CC Desc',
    'Role',
    'Job Title',
    'Entry Date',
    'Actual',
    'Plan'
    ]]



# Import other labor plan for Core Labs

plan_other_cl = pd.read_excel('Plan_hc_2022-23.xlsx', sheet_name='Plan_other_CoreLabs',
                    dtype=object) 

plan_other_cl.head()
plan_other_cl.dtypes

plan_other_cl['Entry Date'].fillna('Not assigned', inplace=True)

hc_data_cl = pd.concat([hc_data_cl, plan_other_cl])

hc_data_cl['Plan Other Labor'].fillna(0, inplace=True)

hc_data_cl['Division_ID'] = hc_data_cl['Division_ID'].astype('int64')
hc_data_cl['Department_ID'] = hc_data_cl['Department_ID'].astype('int64')
hc_data_cl['CostCenter_ID'] = hc_data_cl['CostCenter_ID'].astype('int64')
hc_data_cl['KAUSTID'] = hc_data_cl['KAUSTID'].astype('str')
hc_data_cl['Supervisor_ID'] = hc_data_cl['Supervisor_ID'].astype('str')
hc_data_cl['Actual Position'] = hc_data_cl['Actual Position'].astype('str')
hc_data_cl['Aprroved Position'] = hc_data_cl['Aprroved Position'].astype('str')
hc_data_cl['Approved Classification'] = hc_data_cl['Approved Classification'].astype('str')
hc_data_cl['Approved Grade'] = hc_data_cl['Approved Grade'].astype('str')
hc_data_cl['Position'] = hc_data_cl['Position'].astype('int64')
hc_data_cl['Actual'] = hc_data_cl['Actual'].astype('float64')
hc_data_cl['Plan'] = hc_data_cl['Plan'].astype('float64')
hc_data_cl['Plan Other Labor'] = hc_data_cl['Plan Other Labor'].astype('int64')

hc_data_cl['Approved Grade'] = hc_data_cl['Approved Grade'].replace({"nan": "0"})



# Save Data in Excel

hc_data_cl.to_excel('hc_data_cl_mar_results.xlsx', index = False) # Change excel file name every month
















