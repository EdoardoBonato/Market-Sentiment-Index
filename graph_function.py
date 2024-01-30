# -*- coding: utf-8 -*-
"""
Created on Mon Jan  1 19:28:18 2024

@author: edobo
"""

#import useful packages
import pandas as pd
import numpy as np
import openpyxl
import plotly.express as px
import plotly.io as pio
import plotly.graph_objects as go
import os

def graph_creation(variable = None, balanced = True, weighted = True, area = None):
    graph_data = None
    if area == 'USA' or area == 'United States of America' or area == 'US' or area == 'us' :
        area = 'USA'
    if area == 'EU' or area == 'European Union' or area == 'Europe' or area == 'eu':
        area ='EU'
    if area == None:
        raise ValueError('Please insert a correct area')
    #graph construction with plotly
    if balanced == True :
       path = r'C:\Users\edobo\OneDrive\Desktop\THINGS TO SEND' + '\\' + variable + '_' + area + '_balanced_index.xlsx'
       graph_data = pd.read_excel(path, sheet_name = 'index', header = None, index_col = 0) 
    if balanced == False:
       path = r'C:\Users\edobo\OneDrive\Desktop\THINGS TO SEND' + '\\' + variable + '_' + area + '_unbalanced_index.xlsx'
       graph_data = pd.read_excel(path, sheet_name = 'index', header = None, index_col = 0)
    
    if weighted == True:
        x = list(graph_data.loc['year_weighted'])
        x_rev = list(graph_data.loc['year_weighted'][::-1])
        y_ub =[graph_data.loc['index_mean_weighted'][value] + graph_data.loc['index_std_weighted'][value] for value in range(1,len(graph_data.loc['index_mean_weighted'])+1)]
        y_lb = [graph_data.loc['index_mean_weighted'][value] - graph_data.loc['index_std_weighted'][value] for value in range(1, len(graph_data.loc['index_mean_weighted'])+1)]
        y_lb = y_lb[::-1]
    

        transition_point = 5
        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(
            x = x[transition_point-1:] + x_rev[:transition_point-1],
            y = y_ub[transition_point-1:] + y_lb[:transition_point-1],
            fill = 'toself',fillcolor = 'rgba(0,100,80,0.2)',
            line_color ='rgba(255,255,255,0)', 
            showlegend=False
            ))
   
        if variable == 'cost':
            fig1.add_trace(go.Scatter(
                x = x[:transition_point] + x_rev[:transition_point],
                y = graph_data.loc['index_mean_weighted'][:transition_point],
                line_color='rgb(0,100,80)',
                name = 'cost_actual'
                ))
            fig1.add_trace(go.Scatter(
                x = x[transition_point-1:] + x_rev[transition_point-1:],
                y = graph_data.loc['index_mean_weighted'][transition_point-1:],
                line=dict(color='rgb(0,100,80)', dash='dot'),
                name = 'cost_forecasts'
                ))
            print(x)
            
        if variable == 'revenue':
            fig1.add_trace(go.Scatter(
                x = x[:transition_point] + x_rev[:transition_point],
                y = graph_data.loc['index_mean_weighted'][:transition_point],
                line_color='rgb(0,100,80)',
                name = 'revenue_actual'
                ))
            fig1.add_trace(go.Scatter(
                x = x[transition_point-1:] + x_rev[transition_point-1:],
                y = graph_data.loc['index_mean_weighted'][transition_point-1:],
                line=dict(color='rgb(0,100,80)', dash='dot'),
                name = 'revenue_forecasts'
                ))
            
    if weighted == False:
        x = list(graph_data.loc['year'])
        x_rev = list(graph_data.loc['year'][::-1])
        y_ub =[graph_data.loc['index_mean'][value] + graph_data.loc['index_std'][value] for value in range(1,len(graph_data.loc['index_mean'])+1)]
        y_lb = [graph_data.loc['index_mean'][value] - graph_data.loc['index_std'][value] for value in range(1, len(graph_data.loc['index_mean'])+1)]
        y_lb = y_lb[::-1]
     

        transition_point = 5
        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(
             x = x[transition_point-1:] + x_rev[:transition_point-1],
             y = y_ub[transition_point-1:] + y_lb[:transition_point-1],
             fill = 'toself',fillcolor = 'rgba(0,100,80,0.2)',
             line_color ='rgba(255,255,255,0)', 
             showlegend=False
             ))
    
        if variable == 'cost':
             fig1.add_trace(go.Scatter(
                 x = x[:transition_point] + x_rev[:transition_point],
                 y = graph_data.loc['index_mean'][:transition_point],
                 line_color='rgb(0,100,80)',
                 name = 'cost_actual'
                 ))
             fig1.add_trace(go.Scatter(
                 x = x[transition_point-1:] + x_rev[transition_point-1:],
                 y = graph_data.loc['index_mean'][transition_point-1:],
                 line=dict(color='rgb(0,100,80)', dash='dot'),
                 name = 'cost_forecasts'
                 ))
        if variable == 'revenue':
            fig1.add_trace(go.Scatter(
                     x = x[:transition_point] + x_rev[:transition_point],
                     y = graph_data.loc['index_mean'][:transition_point],
                     line_color='rgb(0,100,80)',
                     name = 'revenue_actual'
                     ))
            fig1.add_trace(go.Scatter(
                     x = x[transition_point-1:] + x_rev[transition_point-1:],
                     y = graph_data.loc['index_mean'][transition_point-1:],
                     line=dict(color='rgb(0,100,80)', dash='dot'),
                     name = 'revenue_forecasts'
                     ))
                 
            fig1.update_xaxes(title_text = "year")
            fig1.update_layout(yaxis = dict( range=[0, 200], dtick = 20))
            print(fig1)
    
    if variable == 'revenue':
        fig1.update_yaxes(title_text = "Index (2018=100)")
    if variable == 'cost':
        fig1.update_yaxes(title_text = "Index (2018=100)" )
        
        
    if balanced == True:
        if weighted == True:
            fig1.update_layout(title_text = "Automotive: " + variable + '_' + area + " ,balanced sample, weighted")
            fig1.write_html(r'C:\Users\edobo\OneDrive\Desktop\THINGS TO SEND'  +'\\' + variable + '_' + area+'_automobiles_balanced_weighted.html')
            return fig1
        if weighted == False:
            fig1.update_layout(title_text = "Automotive: " + variable + '_' + area + " ,balanced sample, unweighted")
            fig1.write_html(r'C:\Users\edobo\OneDrive\Desktop\THINGS TO SEND' + '\\' + variable + '_' + area+ '_automobiles_balanced_unweighted.html')
            return fig1
    if balanced == False:
        if weighted == True:
            fig1.update_layout(title_text = "Automotive: " + variable + '_' + area +  " ,unbalanced sample, weighted")
            fig1.write_html(r'C:\Users\edobo\OneDrive\Desktop\THINGS TO SEND' + '\\' + variable + '_' + area+ '_automobiles_unbalanced_weighted.html')
            return fig1
        if weighted == False:
            fig1.update_layout(title_text = "Automotive: " + variable + '_' + area + " ,unbalanced sample, unweighted")
            fig1.write_html(r'C:\Users\edobo\OneDrive\Desktop\THINGS TO SEND' + '\\' + variable + '_' + area+ '_automobiles_unbalanced_unweighted.html')
            return fig1
