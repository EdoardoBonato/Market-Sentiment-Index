""""
Created on Mon Oct 30 14:00:23 2023

@author: bonated
"""

import pandas as pd
import statistics
import os
import numpy as np
import openpyxl
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#display all in the console
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

#import estimates dataset dataset
path_estimates = r"C:\your_path\estimates_clean.xlsx"
estimates = pd.read_excel(path_estimates , header = 0, na_values = "-")
#list of variables
column_names_estimates = estimates.columns.tolist()
#import historical_data estimates
historical_path = r"C:\your_path\historical_clean.xlsx"
historical = pd.read_excel(historical_path, header = 0, na_values = "-")
historical['FY'] = historical['FY'].astype(int)
column_names_historical = historical.columns.tolist()
historical['EBIT'] = historical['revenue'] - historical['operating_expense']

#see which are the common variables
common_variables = []
for name in column_names_historical:
    if name in column_names_estimates:
      common_variables.append(name)

#we do not want to merge by period. 
common_variables.remove("FY")

#see what are the common companies
#bycode
companies_h = historical["company_code"].unique()
companies_e = estimates["company_code"].unique()
common_companies = []
not_common_companies = []
for company in companies_h:
    if company in companies_e:
        common_companies.append(company)
    else:
        not_common_companies.append(company)

#byname
companiesn_h = historical["company_name"].unique()
companiesn_e = estimates["company_name"].unique()
common_companiesn = []
not_common_companiesn = []
for company in companiesn_h:
    if company in companiesn_e:
        common_companiesn.append(company)
    else:
        not_common_companiesn.append(company)
#they correspond

#---merging operations----#        
#merge the two datasets with keys = common variables
merged = pd.merge(historical, estimates, on = common_variables, how = "outer")
merged_names = merged.columns.tolist()
merged["FY_x"].unique()
#copy the merged dataset, in order to modify it *optional
new_merged = merged.copy()

#get rid of companies which are not in both historical and estimates
group_by_general = new_merged.groupby('company_name')
new_merged = group_by_general.filter(lambda x : len(x) == 9)

new_merged["FY_x"].fillna(merged["FY_y"], inplace = True)
new_merged = new_merged.drop("FY_y", axis = 1)
new_merged = new_merged.rename(columns = {'FY_x': 'FY'})
new_merged['EBIT'].fillna(new_merged['EBIT_mean'], inplace = True)
new_merged = new_merged.drop('EBIT_mean', axis = 1)

#use column variable to store both actual and forecast data, revenue
new_merged['revenue'].fillna(new_merged['revenue_mean'], inplace = True)
new_merged = new_merged.drop("revenue_mean", axis = 1)

#use column variable to store both actual and forecast data, cost
new_merged['cost'].fillna(new_merged['COGS_mean'], inplace = True)
new_merged = new_merged.drop('COGS_mean', axis = 1)
new_merged = new_merged.sort_values(by= ["company_name" , "FY"])
number_companies = new_merged['company_name'].nunique()

path = r"C:\your_path\merged.xlsx"
writer = pd.ExcelWriter(path, engine="xlsxwriter", mode='w')
new_merged.to_excel(writer, sheet_name = 'Data')
writer.close()

'''
#-------------------------------------------------------------------------------------------------------------------#
#------------------------ optional for data info--------------------------------------------------------------------#
#-------------------------------------------------------------------------------------------------------------------#
##PRELIMINARY STATISTICS PART
#missing values 
missing_values = new_merged.nunique()
european_company = new_merged[new_merged.europe == 1]
group_by_eu = european_company.groupby('company_name')
revenuemv_per_company = group_by_eu['revenue'].apply(lambda x : x.isnull().sum())
revenue_mv = [(index, value) for index,value in revenuemv_per_company.items() if value != 0]
costmv_per_company = group_by_eu['cost'].apply(lambda x : x.isnull().sum())
cost_mv = [(index, value) for index,value in costmv_per_company.items() if value != 0]
companies_nr = new_merged['company_name'].nunique()
common_mv = []
for company, value in revenue_mv:
    for companies,value in cost_mv:
        if company == companies:
            common_mv.append(company)
european_company['company_name'].nunique()

#export to excel
path_info = r''
revenue_mv_df = pd.DataFrame(revenue_mv)
cost_mv_df = pd.DataFrame(cost_mv)
writer = pd.ExcelWriter(path_info, engine="xlsxwriter", mode='w')
revenue_mv_df.to_excel(writer, sheet_name = 'revenue_mv')
cost_mv_df.to_excel(writer, sheet_name = 'cost_mv')
writer.close()


'''
















