# -*- coding: utf-8 -*-
"""
Created on Mon Apr  3 08:34:29 2023

@author: londoncm
"""

import pandas as pd
import os

# Adding the path where is the data located

os.chdir(r'Z:\Python\Test\HC_data') 

os.listdir()

# --------------Import HC Data (Excel)-------------------------

df1 = pd.read_excel('hc_apr_2023.xlsx', dtype=object) # Change file name every month

df1.head()
df1.dtypes

# See if Ramzi is in Research Funding

def ramzi_dep(row):
    if row['Employee Name'] ==  'Ramzi Idoughi' and row['RH3 Department'] ==  'Research Funding':    
        row['RH3 Department'] =  'Research Centers'
    else:
       row['RH3 Department'] =  row['RH3 Department']
    return row 
    
df1 = df1.apply(ramzi_dep, axis=1)

def ramzi_cc(row):
    if row['Employee Name'] ==  'Ramzi Idoughi':    
        row['Cost Center'] =  '20410'
    else:
       row['Cost Center'] =  row['Cost Center']
    return row 
    
df1 = df1.apply(ramzi_cc, axis=1)

def ramzi_desc(row):
    if row['Employee Name'] ==  'Ramzi Idoughi':    
        row['CC Desc'] =  'VCC'
    else:
       row['CC Desc'] =  row['CC Desc']
    return row 
    
df1 = df1.apply(ramzi_desc, axis=1)




def damien_dep(row):
    if row['Employee Name'] ==  'Damien James Lightfoot' and row['RH3 Department'] ==  'Research Funding':    
        row['RH3 Department'] =  'Vice President Resea'
    else:
       row['RH3 Department'] =  row['RH3 Department']
    return row 
    
df1 = df1.apply(damien_dep, axis=1)

def damien_cc(row):
    if row['Employee Name'] ==  'Damien James Lightfoot':    
        row['Cost Center'] =  '30023'
    else:
       row['Cost Center'] =  row['Cost Center']
    return row 
    
df1 = df1.apply(damien_cc, axis=1)

def damien_desc(row):
    if row['Employee Name'] ==  'Damien James Lightfoot':    
        row['CC Desc'] =  'Research Support&Val'
    else:
       row['CC Desc'] =  row['CC Desc']
    return row 
    
df1 = df1.apply(damien_desc, axis=1)

def damien_sup(row):
    if row['Employee Name'] ==  'Damien James Lightfoot':    
        row['Supervisor'] =  '100335'
    else:
       row['Supervisor'] =  row['Supervisor']
    return row 
    
df1 = df1.apply(damien_sup, axis=1)

def damien_job(row):
    if row['Employee Name'] ==  'Damien James Lightfoot':    
        row['Job Title'] =  'Strategic Initiatives Manager (SHI)'
    else:
       row['Job Title'] =  row['Job Title']
    return row 
    
df1 = df1.apply(damien_job, axis=1)


# Remove rows when department is Core Lab & Research Funding
df1 = df1.loc[df1['Employee Name'] != 'Eman Mousa A. Alhajji', :]
df1 = df1.loc[df1['RH3 Department'] != 'Core Labs', :]
df1 = df1.loc[df1['RH3 Department'] != 'Research Funding', :]
df1 = df1.loc[df1['RH3 Department'] != 'Divisions and Facult', :]

# Remove colummns we do not use in hc Data

df1 = df1.drop([
    'WBS Element',
    'WBS Desc',
    'Employee Group',
    'Organizational Unit',
    'OU Desc',
    'Master Cost Center',
    'MCC Desc',
    'User Name',
    'Contract End Date',
    ], axis=1)

# Null reeplacement and new columns

df1['RH3 Division'] = df1['RH3 Division'].replace({"Vice President Resea": 3})
df1['RH3 Department'] = df1['RH3 Department'].replace({"Vice President Resea": 1000, "Research Centers": 1010})
df1['Cost Center'] = df1['Cost Center'].replace({"#": 0})
df1['Supervisor'] = df1['Supervisor'].replace({"#": 10000})
df1['Classification'] = df1['Classification'].replace({"PTSA": "DSFTEC", 
                                                       "Post-Doctoral":"FTE/Postdoc",
                                                       "Research":"Research/Engineer",
                                                       "Staff Scientist-Engineer":"Research/Engineer"})

df1['Grade'] = df1['Grade'].replace({"1R": "R1", "2R":"R2", "3R":"R3", "4R":"R4",
                                                       "5R":"R5", "6R":"R6",
                                                       "1P":"P1", "2P":"P2",
                                                       "3P":"P3", "4P":"P4",
                                                       "5P":"P5", "6P":"P6",
                                                       "7P":"P7", "8P":"P8"})


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
    else:
       row['Type Position'] = 'Vacant position'
    return row   


df1 = df1.apply(Type_Position, axis=1)


# Change type of column

df1['Actual'] = df1['Actual'].astype('float64')
df1['Cost Center'] = df1['Cost Center'].astype('int64')



df1.dtypes


# Table with Onboard positions

df2 = df1.loc[df1['Type Position'] == 'Onboard Position', :]

df3 = df1.loc[df1['Type Position'] != 'Onboard Position', :]

# --------------Import Plan HC Data 2022-23 (Excel)-------------------------

plan_data = pd.read_excel('Plan_hc_2022-23.xlsx', sheet_name='plan_hc',
                    dtype=object) 

plan_data.head()
plan_data.dtypes

plan_data['RH3 Division'] = plan_data['RH3 Division'].astype('int64')
plan_data['RH3 Department'] = plan_data['RH3 Department'].astype('int64')
plan_data['Actual'] = plan_data['Actual'].astype('float64')
plan_data['Position'] = plan_data['Position'].astype('str')


# Table with vacant positions

plan_data['Check'] = plan_data.Position.isin(df2.Position)
print(plan_data['Check'].value_counts())

vacants = plan_data.loc[plan_data['Check'] == False, :]

vacants['KAUSTID'].fillna('100000', inplace=True)
vacants['Employee Name'].fillna('Vacant Position', inplace=True)

vacants['Classification'] = vacants['Classification'].replace({"Research":"Research/Engineer",
                                                       "Staff Scientist-Engineer":"Research/Engineer"})

vacants['Entry Date'].fillna('NA', inplace=True)


vacants['Supervisor'].fillna('100000', inplace=True)

vacants = vacants.drop([
    'Check',
    ], axis=1)

vacants = vacants.apply(Type_Position, axis=1)

vacants['Type Position'] = vacants['Type Position'].replace({"Onboard Position": "Vacant Position"})

vacants['Cost Center'] = vacants['Cost Center'].astype('int64')


def approve_grade_vacant(row):
    
    if row['Type Position']=='Vacant Position':    
        row['Approved Grade'] = row['Grade']
    else:
        row['Approved Grade'] = '0'
    return row   


vacants = vacants.apply(approve_grade_vacant, axis=1)

# Identify Wrongly Charged

df2['Check'] = df2.Position.isin(plan_data.Position)
print(df2['Check'].value_counts())

def wrongly(row):
    if row['Check'] ==  False:    
        row['Type Position'] =  'Wrongly Charged'
    else:
       row['Type Position'] =  row['Type Position']
    return row 
    
df2 = df2.apply(wrongly, axis=1)

# Bring grade based on the plan

grade_plan = pd.read_excel('Plan_hc_2022-23.xlsx', sheet_name='grade_plan',
                    dtype=object)

grade_plan['Position'] = grade_plan['Position'] .astype('str')


df2 = pd.merge(df2, grade_plan, on='Position', how='left')

def approved_grade(row):
    if row['Type Position'] ==  'Onboard Position':    
        row['Approved Grade'] =  row['Grade_y']
    else:
       row['Approved Grade'] =  '0'
    return row 

df2 = df2.apply(approved_grade, axis=1)

df2 = df2.drop([
    'Grade_y',
    ], axis=1)


df2 = df2.rename(columns={
    'Grade_x': 'Grade'
    })


df1 = pd.concat([df2, df3])

df1 = df1.drop([
    'Check',
    ], axis=1)



# Add vacant position to hc data

hc_data = pd.concat([df1, vacants])


# Create new columns

def approve_position(row):
    
    if row['Type Position']== 'Onboard Position' or row['Type Position']== 'Vacant Position':    
        row['Aprroved Position'] = row['Position']
    else:
       row['Aprroved Position'] = '0'
    return row   


hc_data = hc_data.apply(approve_position, axis=1)

def actual_position(row):
    
    if row['Type Position']== 'Vacant Position':    
        row['Actual Position'] = '0'
    else:
       row['Actual Position'] = row['Position']
    return row   

hc_data = hc_data.apply(actual_position, axis=1)

def approve_classification(row):
    
    if row['Type Position']== 'Onboard Position' or row['Type Position']== 'Vacant Position':    
        row['Approved Classification'] = row['Classification']
    else:
       row['Approved Classification'] = '0'
    return row   

hc_data = hc_data.apply(approve_classification, axis=1)

def position(row):
    
    if row['Type Position']== 'Onboard Position' or row['Type Position']== 'Vacant Position':    
        row['Position#'] = 1
    elif row['Type Position']== 'Other Labor':
        row['Position#'] = 2
    elif row['Type Position']== 'Wrongly Charged':
        row['Position#'] = 3   
    else:
       row['Position#'] = 4
    return row 

hc_data = hc_data.apply(position, axis=1)



hc_data['CC Desc'] = hc_data['CC Desc'].replace({"Office of A-VPR & CO": "Office of A-VPR & COO",
                                                 "Research Funding&Ser":"Research Funding and Services (RFS)", 
                                                 "Research Support&Val":"Research Support and Valorization (RSV)", 
                                                 "Research Translation":"Research Translation and Partnerships (RTP)",
                                                 "RC Center for Desert":"CDA",
                                                 "Resilient Computing":"RC3",
                                                 "Research Translation and Partnerships (R": 
                                                     "Research Translation and Partnerships (RTP)"})


def esg_reclassified(row):
    
    if row['Type Position']==  'Vacant Position':    
        row['ESGReclassified_ID'] =  '0'
    else:
       row['ESGReclassified_ID'] = row['Grade'] 
    return row 


hc_data = hc_data.apply(esg_reclassified, axis=1)


def plan(row):
    
    if row['Type Position']== 'Onboard Position' or row['Type Position']== 'Vacant Position':     
        row['Plan'] = row['Actual'] 
    else:
       row['Plan'] = 0
    return row 


hc_data = hc_data.apply(plan, axis=1)



# Import table with roles

roles = pd.read_excel('Plan_hc_2022-23.xlsx', sheet_name='Roles',
                     dtype=object) 

roles.head()
roles.dtypes

roles['Position'] = roles['Position'].astype('str')


# Add roles to HC_data

hc_data = pd.merge(hc_data, roles, on='Position', how='left')

hc_data['Role'].fillna('Not Assigned', inplace=True)




# Remove columns

hc_data = hc_data.drop([
    'Grade',
    'Position',
    ], axis=1)

# Rename Columns

hc_data = hc_data.rename(columns={
    'RH3 Division': 'Division_ID',
    'RH3 Department': 'Department_ID',
    'Cost Center': 'CostCenter_ID',
    'Position#':'Position',
    'Supervisor':'Supervisor_ID'
    })


# Change the columns order

hc_data = hc_data[[
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


# Import other labor plan

plan_other = pd.read_excel('Plan_hc_2022-23.xlsx', sheet_name='Plan_other_RC_RO',
                    dtype=object) 

plan_other.head()
plan_other.dtypes

plan_other['Entry Date'].fillna('Not assigned', inplace=True)

hc_data = pd.concat([hc_data, plan_other])

hc_data['Plan Other Labor'].fillna(0, inplace=True)

hc_data['Division_ID'] = hc_data['Division_ID'].astype('int64')
hc_data['Department_ID'] = hc_data['Department_ID'].astype('int64')
hc_data['CostCenter_ID'] = hc_data['CostCenter_ID'].astype('int64')
hc_data['KAUSTID'] = hc_data['KAUSTID'].astype('str')
hc_data['Supervisor_ID'] = hc_data['Supervisor_ID'].astype('str')
hc_data['Actual Position'] = hc_data['Actual Position'].astype('int64')
hc_data['Aprroved Position'] = hc_data['Aprroved Position'].astype('str')
hc_data['Approved Classification'] = hc_data['Approved Classification'].astype('str')
hc_data['Approved Grade'] = hc_data['Approved Grade'].astype('str')
hc_data['Position'] = hc_data['Position'].astype('int64')
hc_data['Actual'] = hc_data['Actual'].astype('float64')
hc_data['Plan'] = hc_data['Plan'].astype('float64')
hc_data['Plan Other Labor'] = hc_data['Plan Other Labor'].astype('int64')

hc_data['Approved Grade'] = hc_data['Approved Grade'].replace({"nan": "0"})



# Save Data in Excel

hc_data.to_excel('hc_data_apr_results.xlsx', index = False) # Change excel file name every month












