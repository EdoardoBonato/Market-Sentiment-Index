# -*- coding: utf-8 -*-
"""
Created on Tue Oct 10 14:24:49 2023

@author: bonated
"""
#importing useful packages
import pandas as pd
import os
import csv
import os
import openpyxl
from openpyxl import load_workbook
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

#import general dataset
path_historical = r'C:\your_path\EC All Yearly Historical Final 05-19-23.xlsb'
historical = pd.read_excel(path_historical, header = 4, na_values = "-")
historical.columns = historical.columns.str.replace("\n", "").str.replace(" ", "").str.replace(r"\(€MM\)", "").str.replace("-", "").str.replace(",", "")

#assign new names to variable
name = {'RICs' : 'company_code', 'Period' : 'FY', 'CompanyCommonName' : 'company_name', 'ISINCode' : 'ISIN_code', 
       'CountryofHeadquarters' : 'HQ_ctry', 'NACEClassification' : 'NACE', 'NAICSNationalIndustryName' : 'NAICS_natl_name',
       'NAICSNationalIndustryCode' : 'NAICS_natl_code', 'NAICSInternationalIndustryName' : 'NAICS_intl_name' , 
       'NAICSInternationalIndustryCode' : 'NAICS_intl_code', 'NAICSSectorName' : 'NAICS_sector_name',
       'NAICSSectorCode' : 'NAICS_sector_code', 'GICSIndustryName' :  'GICS_ind_name', 
       'GICSIndustryCode' : 'GICS_ind_code', 'GICSSubIndustryName' : 'GICS_subind_name', 'GICSSubIndustryCode' : 'GICS_subind_code'}

name2 = {
    'TotalRevenue': 'revenue',
    'NetIncomeAfterTaxes': 'income_after_tax',
    'NetIncomeBeforeTaxes': 'income_before_tax',
    'TotalOperatingExpense': 'operating_expense',
    'CostofRevenueTotal': 'cost',
    'GrossProfit': 'gross_profit',
    'InterestExpenseNetOperating': 'interest_expense_operating',
    'InterestExpenseNetNonOperating': 'interest_expense_non_operating',
    'NonInterestExpenseBank': 'non_interest_expense_banks',
    'GrossMarginPercent': 'gross_margin_pc',
    'OperatingMarginPercent': 'operating_margin_pc',
    'TotalAssetsReported': 'assets',
    'TotalCurrentAssets': 'current_assets',
    'TotalCurrentLiabilities': 'current_liabilities',
    'TotalDebt': 'debt',
    'TotalLongTermDebt': 'lt_debt',
    'TotalLiabilities': 'liabilities',
    'TotalEquity': 'equity',
    'CashandShortTermInvestments': 'cash_st_investment',
    'TotalInventory': 'inventory',
    'TotalInvestmentSecurities': 'investment_securities',
    'NetIncomeMean': 'income_mean',
    'IntangiblesNet': 'intangibles_net',
    'IntangiblesGross': 'intangibles_gross',
    'LongTermInvestments': 'investment_lt',
    'TotalShortTermBorrowings': 'borrowing_st',
    'CommonStockTotal': 'common_stock',
    'Cash': 'cash',
    'CashandEquivalents': 'cash_and_equivalent',
    'ShortTermInvestments': 'investment_st',
    'ResearchAndDevelopment': 'RD',
    'FuelExpense': 'fuel_expenses',
    'InventoriesFinishedGoods': 'inv_finished_goods',
    'InventoriesWorkInProgress': 'inv_wip',
    'InventoriesRawMaterials': 'inv_rawmat',
    'InventoriesOther': 'inv_other',
    'Inventories(CF)': 'inv_CF',
    'LaborAndRelatedExpense': 'labour_expenses',
    'EnergyUseTotal(Gigajoules)': 'energy_use_GJ',
    'EnergyPurchasedDirect(Gigajoules)': 'energy_purchsd_GJ',
    'IndirectEnergyUse(Gigajoules)': 'indirect_energy_use_GJ',
    'EnergyProducedDirect(Gigajoules)': 'energy_produced_GJ',
    'ElectricityPurchased(Gigajoules)': 'electricity_purchsd_GJ',
    'ElectricityProduced(Gigajoules)': 'electricity_producd_GJ',
    'RenewableEnergyPurchased(Gigajoules)': 'res_purchsd_GJ',
    'RenewableEnergyProduced(Gigajoules)': 'res_producd_GJ',
    'RenewableEnergyUse': 'res_use',
    'GreenBuildings': 'green_buildings',
    'EnvironmentalSupplyChainManagement': 'env_supply_chain_mgt',
    'EnvironmentalSupplyChainMonitoring': 'env_supply_chain_monitoring',
    'EnvironmentalControversiesCount': 'env_controversies',
    'TotalRenewableEnergy(Gigajoules)': 'res_total',
    'TotalRenewableEnergyToEnergyUseinmillion(Gigajoules)': 'res_to_total_energy',
    'CO2EquivalentEmissionsTotal(Tonnes)': 'CO2eq_emission',
    'CO2EstimationMethod': 'CO2_estimation',
    'TotalCO2EquivalentEmissionsToMillionRevenuesUSDYoY(Percent)': 'CO2eq_to_revenue_USD_pc_yoy',
    'CO2EquivalentEmissionTotalYoY(Percent)': 'CO2eq_pc_yoy',
    'EnvironmentalProducts': 'env_products',
    'EnvironmentalRDExpenditures': 'env_RD',
    'EcoDesignProducts': 'eco_design_product',
    'Renewable/CleanEnergyProducts': 'res_clean_product',
    'GreenCapex': 'green_capex',
    'GreenCapexTarget': 'green_capex_target',
    'ESGScoreGrade': 'ESG_score',
    'PolicySkillsTraining': 'skills_training_policy',
    'PolicyForcedLabor': 'forced_labor_policy',
    'PolicyHumanRights': 'human_rights_policy',
    'SalaryGap': 'salary_gap',
    'SalariesandWagesfromCSRreporting': 'salaries_from_CSR',
    'NumberofEmployees': 'employees'
}

historical = historical.rename(columns = name)
historical = historical.rename(columns = name2)
historical = historical.dropna(subset=['FY'])

#create new europe variable
european_countries = ["France","Denmark","Netherlands","Ireland; Republic of","Germany","Belgium", "Spain","Sweden","Italy","Finland","Czech Republic","Luxembourg","Portugal","Poland","Austria","Hungary","Romania","Greece","Slovenia","Cyprus","Malta"]
historical["europe"] = historical["HQ_ctry"].apply(lambda x : 1 if x in european_countries else 0)
historical[['NACE1', 'NACE2']] = historical['NACE'].str.split(" \(NACE\) \(", expand=True)
historical['NACE2'] = historical['NACE2'].str.replace(")", "")
historical = historical.rename(columns={'NACE1': 'NACE_name', 'NACE2': 'NACE_code'})
country_mapping = {'United States of America' : 'USA'}
historical['HQ_ctry'] = historical['HQ_ctry'].map(country_mapping)
for year in ["2017", "2018", "2019", "2020", "2021", "2022"]:
    historical["FY"].replace(to_replace=f"FY{year}", value=year, regex=True, inplace=True)
    
# Drop the original 'NACE' column
historical = historical.drop('NACE', axis=1)

#create a list for columns(variables)names
column_names = historical.columns.tolist()

path_historical_clean = r"C:\your_path\historical_clean.xlsx"

writer = pd.ExcelWriter(path_historical_clean, engine="xlsxwriter", mode='w')
historical.to_excel(writer, sheet_name = "Data")
writer.close()

#------------------------------------------------------------------------------------------------------------------#
#------------------------------------------------optional for data info--------------------------------------------#
#------------------------------------------------------------------------------------------------------------------#
'''
#missing values
missing_values_count = historical.isnull().sum()
missing_values_count_percentage = ((missing_values_count / 9837)*100)
#per company
european_company = historical[historical.europe == 1]
group_by = european_company.groupby('company_name')
per_column = group_by['revenue'].apply(lambda x : x.isnull().sum())
companies_mv = [(index, value) for index,value in per_column.items() if value != 0]
 
#check whether there are OUTLIERS( high percentage changeyear by year for the same company_name)
group_by = historical.groupby("company_name")
numeric_variables = "revenue income_after_tax income_before_tax operating_expense cost gross_profit interest_expense_operating interest_expense_non_operating non_interest_expense_banks gross_margin_pc operating_margin_pc assets current_assets current_liabilities debt lt_debt liabilities equity cash_st_investment inventory investment_securities income_mean intangibles_net intangibles_gross investment_lt borrowing_st common_stock cash cash_and_equivalent investment_st RD fuel_expenses inv_finished_goods inv_wip inv_rawmat inv_other inv_CF labour_expenses energy_use_GJ energy_purchsd_GJ indirect_energy_use_GJ energy_produced_GJ electricity_purchsd_GJ electricity_producd_GJ res_purchsd_GJ res_producd_GJ res_total res_to_total_energy CO2eq_emission salary_gap salaries_from_CSR employees"
numeric_variables = numeric_variables.split()
outliers_list = {}
for variable in numeric_variables:
    outliers_list[variable] = []
    changes_by_variable = group_by[variable].pct_change()
    n = 0
    for values in changes_by_variable:
        n = n + 1
        #the magnitude of the change is settled to 100x
        if values >= 100 or values <= -100:
            values = (values, n)
            outliers_list[variable].append(values)

#list of outliers in percentage
outliers_percentage = []
for key, size in outliers_list.items():
   tup = (key, len(size),  (len(size)/59238)*100, (len(size)/9837)*100)
   outliers_percentage.append(tup)

#check negative values
non_negative_variables = ["revenue", "electricity_purchsd_GJ", "CO2eq_emission"]
negative_values = []
for element in non_negative_variables:
    for values in historical[element]:
        if values < 0:
            print(values)
            negative_values.append(values)
    
#signal 0 values       
zero_list = {}
for variable in historical.columns:
    zero_list[variable] = []
    n = 0
    for values in historical[variable]:
        n = n + 1
        if values == 0:
            values = (values, n)
            print(values)
            zero_list[variable].append(values)

zero_list_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in zero_list.items()]))
zero_list_df = zero_list_df.dropna(axis = 'columns', how= 'all')

#export all to excel
path_info = r''
writer = pd.ExcelWriter(path_info, engine="xlsxwriter", mode='w')
missing_values_count.to_excel(writer, sheet_name = "missing_values")
zero_list_df.to_excel(writer, sheet_name = 'zero_values')
negative_values = pd.DataFrame(negative_values)
outliers_percentage = pd.DataFrame(outliers_percentage)
negative_values.to_excel(writer, sheet_name = 'negative_values')
outliers_percentage.to_excel(writer, sheet_name = 'outliers_prc')
missing_values_count_percentage.to_excel(writer, sheet_name = 'missing_values_prc')
writer.close()
'''