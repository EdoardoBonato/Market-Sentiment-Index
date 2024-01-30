# -*- coding: utf-8 -*-
"""
Created on Mon Nov  6 17:25:32 2023

@author: bonated
"""
import pandas as pd
import statistics
import os
import numpy as np
import openpyxl
import networkx as nx
import itertools
from itertools import combinations
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

def correspondence_code(code1, code2):
    path_merged = r"merged_historical_estimates.xlsx"
    #import general dataset
    final = pd.read_excel(path_merged, header = 0, na_values = "-")
    correspondence = {}
    cross_tab = pd.crosstab(final[code1], final[code2])
    for key,row in cross_tab.iterrows():
        sector = [col for col,value in row.items() if value != 0]
        correspondence[key] = sector
    return correspondence
