# -*- coding: utf-8 -*-
"""
Created on Wed May 17 11:14:27 2023

@author: londoncm
"""

import pandas as pd
import os

# Adding the path where is the data located

os.chdir(r'Z:\Python\Test\HC_data') 

os.listdir()

# --------------Import HC Data (Excel)-------------------------

df_rf1 = pd.read_excel('hc_funding_apr_2023.xlsx', dtype=object) # Change file name every month

df_rf1.head()
df_rf1.dtypes

# Null reeplacement and new columns

df_rf1['WBS_Masking'] = df_rf1['WBS Element'].str[0:3]


def department(row):
    
    if row['WBS_Masking']== 'BAS':    
        row['RH3 Department'] = 'Baseline Funding'
    elif row['WBS_Masking']== 'DIS':    
        row['RH3 Department'] = 'Discretionary Fund'
    elif row['WBS_Masking']== 'FCC':    
        row['RH3 Department'] = 'Center Competitive Fund'
    elif row['WBS_Masking']== 'FCS':    
        row['RH3 Department'] = 'Center Partnership Fund'
    elif row['WBS_Masking']== 'GIF':    
        row['RH3 Department'] = 'Gift and Donations'
    elif row['WBS_Masking']== 'REI':    
        row['RH3 Department'] = 'President Strategic Initiative'
    elif row['WBS_Masking']== 'REP':    
        row['RH3 Department'] = 'Research Partnership'
    elif row['WBS_Masking']== 'RGC':    
        row['RH3 Department'] = 'External Research'
    elif row['WBS_Masking']== 'URF':    
        row['RH3 Department'] = 'CRGs'
    else:
       row['RH3 Department'] = ''
    return row  


df_rf1 = df_rf1.apply(department, axis=1)


# Remove colummns we do not use in hc Data Res. Funding

df_rf1 = df_rf1.drop([
    'RH2 Department',
    'Cost Center',
    'Expiry Date',
    'Organizational Unit',
    'Exit Date',
    'Begin date: WPBP',
    ], axis=1)


# Null reeplacement

df_rf1['Master Cost Center'] = df_rf1['Master Cost Center'].replace({"#": 0})
df_rf1['Supervisor'] = df_rf1['Supervisor'].replace({"#": 100000})
df_rf1['RH3 Division'] = df_rf1['RH3 Division'].replace({"Vice President Resea": 'Vice President Research'})
df_rf1['Contract End Date'] = df_rf1['Contract End Date'].replace({"#": '9.9.9999'})


# --------------Import Mapping to Faculty data (Excel)-------------------------

mapping_1 = pd.read_excel('Mapping_Research_Fundings.xlsx', sheet_name='Mapping_1',
                    dtype=object) 
mapping_1.dtypes


mapping_1['Faculty KAUST ID'] = mapping_1['Faculty KAUST ID'].astype('int64')



mapping_2 = pd.read_excel('Mapping_Research_Fundings.xlsx', sheet_name='Mapping_2',
                    dtype=object) 

mapping_2.dtypes

mapping_2['Faculty KAUST ID'] = mapping_2['Faculty KAUST ID'].astype('int64')

mapping_2['Role'] = mapping_2['Role'].replace({"AMPMC": 'AMPM'})
mapping_2['Role'] = mapping_2['Role'].replace({'CLI': 'Non-Affiliated'})
mapping_2['Role'] = mapping_2['Role'].replace({'KAUST Artificial Intelligence Initiative': 'Non-Affiliated'})
mapping_2['Role'] = mapping_2['Role'].replace({'KAUST Smart Health Initiative': 'Non-Affiliated'})




# Bring the Faculty KAUST ID

hc_funding = pd.merge(df_rf1, mapping_1, on='WBS Element', how='left')


# Bring the Faculty Name and Role

hc_funding = pd.merge(hc_funding, mapping_2, on='Faculty KAUST ID', how='left')

hc_funding['Role'].fillna('Non-Affiliated', inplace=True)

# Rename Columns

hc_funding = hc_funding.rename(columns={
    'User Name': 'Reporting Manager Name',
    'Contract End Date': 'Contract Ending',
    'Actual':'HC %'
    })


# Change the columns order

hc_funding = hc_funding[[
    'RH3 Division',
    'RH3 Department',
    'WBS Element',
    'WBS Desc',
    'Employee Group',
    'Classification',
    'ESG Reclassified',
    'KAUSTID',
    'Employee Name',
    'Position',
    'Job Title',
    'Job',
    'Master Cost Center',
    'M CC Desc',
    'Entry Date',
    'Supervisor',
    'Reporting Manager Name',
    'HC %',
    'Faculty Name',
    'Faculty KAUST ID',
    'WBS_Masking',
    'Role',
    'Contract Ending',
    ]]



# Save Data in Excel

hc_funding.to_excel('hc_funding_apr_results.xlsx', index = False) # Change excel file name every month


