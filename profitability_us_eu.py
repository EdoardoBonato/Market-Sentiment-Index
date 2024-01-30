# -*- coding: utf-8 -*-
"""
Created on Fri Dec  1 17:24:25 2023

@author: bonated
"""

import pandas as pd
import numpy as np
import openpyxl
import plotly.express as px
import plotly.io as pio
import plotly.graph_objects as go
import os


path_us = r"C:\your_path\revenue_USA_balanced_index.xlsx"
path_eu = r"C:\yout_path\revenue_EU_balanced_index.xlsx"

us = pd.read_excel(path_us,sheet_name = 'index', header = 0)
eu = pd.read_excel(path_eu, sheet_name = 'index', header = 0)
index_us = pd.read_excel(path_us, sheet_name = 'index',header = None, index_col = 0)
index_eu = pd.read_excel(path_eu, sheet_name = 'index', header = None, index_col = 0)

#us
index_mean_titles = [2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025]
x = list(index_us.loc['year'])
x_rev = list(index_us.loc['year'][::-1])
y_ub_us = [index_us.loc['index_mean_weighted'][value] + index_us.loc['index_std_weighted'][value] for value in range(1,len(index_us.loc['index_mean_weighted'])+1)]
y_lb_us = [index_us.loc['index_mean_weighted'][value] - index_us.loc['index_std_weighted'][value] for value in range(1, len(index_us.loc['index_mean_weighted'])+1)]
y_lb_us = y_lb_us[::-1]

transition_point = 5
fig1 = go.Figure()

fig1.add_trace(go.Scatter(
    x = x[transition_point-1:] + x_rev[:transition_point-1],
    y = y_ub_us[transition_point-1:] + y_lb_us[:transition_point-1],
    fill = 'toself',fillcolor = 'rgba(0,100,80,0.2)',
    line_color ='rgba(255,255,255,0)', 
    showlegend=False
    ))

fig1.add_trace(go.Scatter(
   x = x[:transition_point] + x_rev[:transition_point],
   y = index_us.loc['index_mean_weighted'][:transition_point],
   line_color='rgb(0,100,80)',
   name = 'us_actual'
   ))
fig1.add_trace(go.Scatter(
    x = x[transition_point-1:] + x_rev[transition_point-1:],
    y = index_us.loc['index_mean_weighted'][transition_point-1:],
    line=dict(color='rgb(0,100,80)', dash='dot'),
    name = 'us_forecasts'
   ))

#eu
y_ub_eu = [index_eu.loc['index_mean_weighted'][value] + index_eu.loc['index_std_weighted'][value] for value in range(1, len(index_eu.loc['index_mean_weighted'])+1)]
y_lb_eu = [index_eu.loc['index_mean_weighted'][value] - index_eu.loc['index_std_weighted'][value] for value in range(1, len(index_eu.loc['index_mean_weighted'])+1)]
y_lb_eu = y_lb_eu[::-1]


fig1.add_trace(go.Scatter(
    x = x[transition_point-1:] + x_rev[:transition_point-1],
    y = y_ub_eu[transition_point-1:] + y_lb_eu[:transition_point-1],
    fill = 'toself', fillcolor='rgba(255, 0, 0, 0.2)',
    line=dict(color='rgba(255, 255, 255, 0)'), 
    showlegend=False
    ))

fig1.add_trace(go.Scatter(
   x = x[:transition_point] + x_rev[:transition_point],
   y = index_eu.loc['index_mean_weighted'][:transition_point],
   line_color='rgb(255,0,0,1)',
   name = 'eu_actual'
   ))
fig1.add_trace(go.Scatter(
    x = x[transition_point-1:] + x_rev[transition_point-1:],
    y = index_eu.loc['index_mean_weighted'][transition_point-1:],
    line=dict(color ='rgba(255, 0, 0, 0.5)', dash='dot'),
    name = 'eu_forecasts'
   ))

fig1.update_xaxes(title_text = "year")
fig1.update_yaxes(title_text = "Index (2018=100)")
fig1.update_layout(title_text = "Automotive profitability: us and eu, balanced sample, weighted")
fig1.update_layout(yaxis = dict( range=[0, 150], dtick = 20))
fig1.write_html(r'C:\your_path\automobiles_profitability_EBIT_eu_us_without_out.html')

