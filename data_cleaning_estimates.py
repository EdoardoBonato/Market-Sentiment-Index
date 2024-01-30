# -*- coding: utf-8 -*-
"""
Created on Thu Oct 19 09:50:13 2023

@author: bonated
"""

#importing useful packages
import pandas as pd
import os
import openpyxl
from openpyxl import load_workbook
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#display all in the console
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

#import general dataset
path_estimates = r"C:\your_path\EC Estimates Yearly Final 05-26-23.xlsb"
estimates = pd.read_excel(path_estimates, header = 4, na_values = "-")

#get rid of empty rows
estimates = estimates.dropna(subset=['Period'])

#rename variables
estimates.columns = estimates.columns.str.replace("\n", "").str.replace(" ", "").str.replace(r"\(€MM\)", "")



name = {'RICs' : 'company_code', 'Period' : 'FY', 'CompanyCommonName' : 'company_name', 'ISINCode' : 'ISIN_code', 
       'CountryofHeadquarters' : 'HQ_ctry', 'NACEClassification' : 'NACE', 'NAICSNationalIndustryName' : 'NAICS_natl_name',
       'NAICSNationalIndustryCode' : 'NAICS_natl_code', 'NAICSInternationalIndustryName' : 'NAICS_intl_name' , 
       'NAICSInternationalIndustryCode' : 'NAICS_intl_code', 'NAICSSectorName' : 'NAICS_sector_name',
       'NAICSSectorCode' : 'NAICS_sector_code', 'GICSIndustryName' :  'GICS_ind_name', 
       'GICSIndustryCode' : 'GICS_ind_code', 'GICSSubIndustryName' : 'GICS_subind_name', 'GICSSubIndustryCode' : 'GICS_subind_code'}

estimates = estimates.rename(columns =  name)

name2 = {'RevenueMean' : 'revenue_mean', 'NetIncomeMean' : 'income_mean', 'EBITMean' : 'EBIT_mean',
         'OperatingExpMean' : 'operating_exp_mean', 'COGSMean' : 'COGS_mean', 'GrossIncomeMean' : 'gross_income_mean',
         'IntExpMean' : 'int_exp_mean', 'NonInterestExpMean' : 'non_int_exp_mean', 'GPMMean' : 'GPM_mean', 
         'TotalAssetsMean' : 'assets_mean', 'CurrentAssetsMean' : 'current_assets_mean',
         'CurrentLiabilitiesMean' : 'current_liabilities_mean','TotalDebtMean' : 'debt_mean', 
         'TotalLiabilitiesMean' : 'liabilities_mean', 'ShareholdersEquityMean' : 'shareholders_equity_mean',
         'CashEquivalentsMean' : 'cash_equivalents_mean', 'InventoryMean' : 'inventory_mean',
         'NetInvestmentIncomeMean' : 'investment_income_mean', 'NetProfitMean' : 'profit_mean', 
         'NumberofSharesOutstandingMean' : 'nr_shares_oustanding_mean', 'CashEquivalentsMean.1' : 'cash_equivalents_mean', 
         'R&DExpenseMean' : 'R&D_exp_mean', 'InventoryMean.1' : 'inventory_mean', 'TotalCompensationExpenseMean' : 'compensation_exp_mean',
         'RevenueStdDev' : 'revenue_std', 'NetIncomeStdDev' : 'income_std', 'EBITStDev' : 'EBIT_std', 'OperatingExpenseStdDev' : 'operating_exp_std',
         'COGSStdDev' : 'COGS_std', 'GrossIncomeStdDev' : 'gross_income_std', 'InterestExpenseStdDev' : 'int_exp_std',
         'NonInterestExpenseStdDev' : 'non_int_exp_std', 'GrossprofitMarginStdDev(%)' : 'GPM_std_prc', 
         'TotalAssetsStdDev' : 'assets_std', 'CurrentAssetsStdDev' : 'current_assets_std', 'CurrentLiabilitiesStdDev' : 'current_liabilities_std',
         'TotalDebtStdDev': 'debt_std', 'TotalLiabilitiesStdDev' : 'liabilities_std', 'ShareholdersEquityStdDev' : 'shareholders_equity_std',
         'CashEquivalentsStdDev': 'cash_equivalents_std', 'InventoryStdDev' : 'inventory-std', 'NetInvestmentIncomeStdDev' :'investment_income_std',
         'NetProfitStdDev' : 'profit_std', 'NumberofSharesOutstandingStdDev(Shares)' : 'nr_shares_oustanding_std', 'CashEquivalentsStdDev.1' : 'cash_equivalents_std',
         'R&DExpenseStdDev' : 'R&D_exp_std', 'InventoryStdDev.1' : 'inventory_std', 'TotalCompensationExpenseStdDev' : 'compensation_exp_std',
         'RevenueSmartEst' : 'revenue_sest', 'NetprofitSmartEst' : 'profit_sest', 'EBITSmartEst' : 'EBIT_sest', 
         'OperatingExpSmartEst' : 'operating_exp_sest', 'COGSSmartEst' : 'COGS_sest', ';GrossIncomeSmartEst' : 'gross_income_sest',
         'IntExpSmartEst' : 'int_exp_sest', 'NonIntExpSmartEst' : 'non_int_exp_sest', 'GrossProfitmarginSmartEst(%)' : 'gross_profit_margin_sest',
         'TotalAssetsSmartEst' : 'assets_sest', 'NetDebtSmartEst' : 'net_debt_sest', 'ShareholdersEquitySmartEst' : 'shareholders_equity_sest',
         'InventorySmartEst' : 'inventory_sest', 'NetInvestmentIncomeSmartEst' : 'investiment_income_sest', 'NetprofitSmartEst.1' : 'profit_sest',
         'NumberofSharesOutstandingSmartEst(Shares)' : 'nr_shares_outstanding_sest', 'CashEquivalentsSmartEst' : 'cash_equivalents_sest',
         'R&DExpenseSmartEst' : 'R&D_exp_sest', 'InventorySmartEst.1' : 'inventory_sest', 'TotalCompensationExpenseSmartEst' : 'compensation_exp_sest'
         }
estimates = estimates.rename(columns = name2)

#create new europe variable
european_countries = ["France","Denmark","Netherlands","Ireland; Republic of","Germany","Belgium", "Spain","Sweden","Italy","Finland","Czech Republic","Luxembourg","Portugal","Poland","Austria","Hungary","Romania","Greece","Slovenia","Cyprus","Malta"]
estimates["europe"] = estimates["HQ_ctry"].apply(lambda x : 1 if x in european_countries else 0)
# Split the 'NACE' column into two new columns based on the pattern " (NACE) ("
estimates[['NACE1', 'NACE2']] = estimates['NACE'].str.split(" \(NACE\) \(", expand=True)

country_mapping = {'United States of America' : 'USA'}
estimates['HQ_ctry'] = estimates['HQ_ctry'].map(country_mapping)
# Remove ")" from the 'NACE2' column
estimates['NACE2'] = estimates['NACE2'].str.replace(")", "")

# Rename columns
estimates = estimates.rename(columns={'NACE1': 'NACE_name', 'NACE2': 'NACE_code'})

# Drop the original 'NACE' column
estimates = estimates.drop('NACE', axis=1)

#replace FY20xx with xx
for year in ["2023", "2024", "2025"]:
    estimates["FY"].replace(to_replace=f"FY{year}", value=year, regex=True, inplace=True)


#create a list for columns(variables)names
column_names = estimates.columns.tolist()
path_estimates_clean = r"C:\your_path\estimates_clean.xlsx"
writer = pd.ExcelWriter(path_estimates_clean, engine="xlsxwriter", mode='w')
estimates.to_excel(writer, sheet_name = "Data")
writer.close()

#----------------------------------------------------------------------------------------------------------------------------------#
#------------------optional for info on data---------------------------------------------------------------------------------------#
#----------------------------------------------------------------------------------------------------------------------------------#
'''
zero_list = {}
#signal 0 values
for variable in estimates.columns:
    zero_list[variable] = []
    n = 0
    for values in estimates[variable]:
        n = n + 1
        if values == 0:
            values = (values, n)
            zero_list[variable].append(values)
            
zero_list_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in zero_list.items()]))
zero_list_df = zero_list_df.dropna(axis = 'columns', how= 'all')
zero_list_df = zero_list_df.drop(columns = "europe")

#create a list for columns(variables)names
column_names = estimates.columns.tolist()
#list of all missing values
missing_values_count = estimates.isnull().sum()
european_company = estimates[estimates.europe == 1]
zero_list_eu = {}
for variable in european_company.columns:
    zero_list_eu[variable] = []
    n = 0
    for values in european_company[variable]:
        n = n + 1
        if values == 0:
            values = (values, n)
            zero_list_eu[variable].append(values)
    
group_by = european_company.groupby('company_name')
per_column_mv = group_by['revenue_mean'].apply(lambda x : x.isnull().sum())
companies_mv = [(index, value) for index,value in per_column_mv.items() if value != 0]
#standard deviation zero values 
for company, group in group_by['revenue_std']:
    for value in group:
        if value == 0:
            print(company)

#name  of companies which have zero values in the standard deviation
zero_values_company_nr = [company for company, group in group_by if (group['revenue_std'] == 0).any()]

#export all to excel
path_info = r''
writer = pd.ExcelWriter(path_info, engine="xlsxwriter", mode='w')
missing_values_count.to_excel(writer, sheet_name = "missing_values")
zero_list_df.to_excel(writer, sheet_name = 'zero_values')
writer.close()


'''

    