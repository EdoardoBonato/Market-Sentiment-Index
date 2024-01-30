# -*- coding: utf-8 -*-
"""
Created on Tue Nov 28 14:58:00 2023

@author: bonated
"""

#import useful packages
import pandas as pd
import numpy as np
import openpyxl
import plotly.express as px
import plotly.io as pio
import plotly.graph_objects as go
import os
#---personal modules
import sys
sys.path.append(r"C:\Users\edobo\OneDrive\Desktop\THINGS TO SEND")
from correspondence_codes import correspondence_code
from index_function import index_creation
from graph_function import graph_creation
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

#import merged dataset
path_merged = r"C:\Users\edobo\OneDrive\Desktop\THINGS TO SEND\input\merged.xlsx"
final = pd.read_excel(path_merged, sheet_name = 'Data', header = 0, na_values = "-")
#list of variables to list
column_names = final.columns.tolist()

#use the index function. This produces the final output : Excel Datasets and Graphs
index_creation(data = final, variable = 'revenue', area = 'us')
graph_creation(variable = 'revenue',  area = 'us')

index_creation(data = final, variable = 'revenue', area = 'eu')
graph_creation(variable  = 'revenue', area = 'eu')

index_creation(data = final, variable = 'cost', area = 'eu')
graph_creation(variable = 'cost', area = 'eu')

index_creation(data = final, variable = 'cost',  area = 'us')
graph_creation(variable = 'cost',  area = 'us')


