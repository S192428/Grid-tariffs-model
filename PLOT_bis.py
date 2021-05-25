# -*- coding: utf-8 -*-
"""
Created on Tue Mar 30 08:48:01 2021

@author: jeand
"""
from math import *
import csv
import gurobipy as gp
from gurobipy import GRB
import numpy as np
import pandas as pd
from pyomo.environ import *
from random import *
import plotly
import plotly.graph_objs as go
from plotly.offline import download_plotlyjs, init_notebook_mode, plot
from numpy import *
import plotly.express as px

# ----------------------
# Import of data
# ----------------------
# Folder = 'D:/DTU/2A/MSC Thesis/Model/Figure ADJ, 0.05, 5/'
Folder = 'D:/DTU/2A/MSC Thesis/Model/Figure ADJ, 0.05, 3/'
# Folder = 'D:/DTU/2A/MSC Thesis/Model/Figure BC, CO2_0.1/'
# Folder = 'D:/DTU/2A/MSC Thesis/Model/Figure BC, CO2_0.04/'
# Folder = 'D:/DTU/2A/MSC Thesis/Model/'
# Folder = 'D:/DTU/2A/MSC Thesis/Model/Figure BC, Prop_0.75/'
# Folder = 'D:/DTU/2A/MSC Thesis/Model/Figure BC, Prop_0.25/'
# Folder = 'D:/DTU/2A/MSC Thesis/Model/Figure BC, SC_202/'
# Folder = 'D:/DTU/2A/MSC Thesis/Model/Figure BC, SC_606/'
# Folder = 'D:/DTU/2A/MSC Thesis/Model/Right profile/BC/'
# Folder = 'D:/DTU/2A/MSC Thesis/Model/Right profile/BC 2/'

Add_name='Sens_CPT_3_'

All_NM_res = pd.read_excel(Folder + 'Results_compiled.xlsx',header=0,sheet_name='All NM')
All_NNM_res = pd.read_excel(Folder +'Results_compiled.xlsx',header=0,sheet_name='All NNM')

x=[All_NM_res.loc[i,'Volumetric Share'] for i in range(len(All_NM_res))]
y=[All_NM_res.loc[i,'Capacity Share'] for i in range(len(All_NM_res))]

x = list(set(x))
y = list(set(y))
x.sort()
y.sort()

x_axis = zeros(len(x))
y_axis = zeros(len(y))
for i in range(len(x)):
    x_axis[i] = round(x[i],2)
    y_axis[i] = round(y[i],2)

z_axis_fairness = [0,0.1,0.2,0.3,0.4,0.5]
    

fairnessNM = pd.DataFrame(index=x,columns=y)
fairnessNNM = pd.DataFrame(index=x,columns=y)
taxNM = pd.DataFrame(index=x,columns=y)
taxNNM = pd.DataFrame(index=x,columns=y)
carboneNM = pd.DataFrame(index=x,columns=y)
carboneNNM = pd.DataFrame(index=x,columns=y)
totNM = pd.DataFrame(index=x,columns=y)
totNNM = pd.DataFrame(index=x,columns=y)
PVNM = pd.DataFrame(index=x,columns=y)
PVNNM = pd.DataFrame(index=x,columns=y)
BatNM = pd.DataFrame(index=x,columns=y)
BatNNM = pd.DataFrame(index=x,columns=y)
P2HNM = pd.DataFrame(index=x,columns=y)
P2HNNM = pd.DataFrame(index=x,columns=y)
HSNM = pd.DataFrame(index=x,columns=y)
HSNNM = pd.DataFrame(index=x,columns=y)
PLNM = pd.DataFrame(index=x,columns=y)
PLNNM = pd.DataFrame(index=x,columns=y)
PROtotNM = pd.DataFrame(index=x,columns=y)
PROtotNNM = pd.DataFrame(index=x,columns=y)
PAStotNM = pd.DataFrame(index=x,columns=y)
PAStotNNM = pd.DataFrame(index=x,columns=y)
VNPNM = pd.DataFrame(index=x,columns=y)
VNPNNM = pd.DataFrame(index=x,columns=y)
CNPNM = pd.DataFrame(index=x,columns=y)
CNPNNM = pd.DataFrame(index=x,columns=y)
FNPNM = pd.DataFrame(index=x,columns=y)
FNPNNM = pd.DataFrame(index=x,columns=y)
for i in range(len(All_NM_res)):
    alpha = All_NM_res.loc[i,'Volumetric Share']
    beta = All_NM_res.loc[i,'Capacity Share']
    fairnessNM.loc[alpha,beta] = float(All_NM_res.loc[i,'fairness'])
    fairnessNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'fairness'])
    taxNM.loc[alpha,beta] = float(All_NM_res.loc[i,'Tax avoidance'])
    taxNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'Tax avoidance'])
    carboneNM.loc[alpha,beta] = float(All_NM_res.loc[i,'CO2 avoided'])
    carboneNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'CO2 avoided'])
    totNM.loc[alpha,beta] = float(All_NM_res.loc[i,'Total Cost of the network'])
    totNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'Total Cost of the network'])
    PVNM.loc[alpha,beta] = float(All_NM_res.loc[i,'PV'])
    PVNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'PV'])
    BatNM.loc[alpha,beta] = float(All_NM_res.loc[i,'Battery'])
    BatNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'Battery'])
    P2HNM.loc[alpha,beta] = float(All_NM_res.loc[i,'P2H'])
    P2HNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'P2H'])
    HSNM.loc[alpha,beta] = float(All_NM_res.loc[i,'HS'])
    HSNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'HS'])
    PLNM.loc[alpha,beta] = float(All_NM_res.loc[i,'Peak Load'])
    PLNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'Peak Load'])
    PROtotNM.loc[alpha,beta] = float(All_NM_res.loc[i,'Total Cost Prosumer'])
    PROtotNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'Total Cost Prosumer'])
    PAStotNM.loc[alpha,beta] = float(All_NM_res.loc[i,'Total Cost Passive'])
    PAStotNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'Total Cost Passive'])
    VNPNM.loc[alpha,beta] = float(All_NM_res.loc[i,'VNP'])
    VNPNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'VNP'])
    CNPNM.loc[alpha,beta] = float(All_NM_res.loc[i,'CNP'])
    CNPNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'CNP'])
    FNPNM.loc[alpha,beta] = float(All_NM_res.loc[i,'FNP'])
    FNPNNM.loc[alpha,beta] = float(All_NNM_res.loc[i,'FNP'])
    
# for col in fairnessNM:
#     fairnessNM[col] = pd.to_numeric(fairnessNM[col], errors='coerce')
# for col in fairnessNNM:
#     fairnessNNM[col] = pd.to_numeric(fairnessNNM[col], errors='coerce')

# counter = 0
# while counter < len(x)-3:
#     fairnessNM[x[counter]] = fairnessNM[x[counter]].interpolate(method = 'cubic',axis = 0,limit_direction = 'both')
#     fairnessNM.loc[y[counter]] = fairnessNM.loc[y[counter]].interpolate(method = 'cubic',axis = 0,limit_direction = 'both')
#     fairnessNNM[x[counter]] = fairnessNNM[x[counter]].interpolate(method = 'cubic',axis = 0,limit_direction = 'both')
#     fairnessNNM.loc[y[counter]] = fairnessNNM.loc[y[counter]].interpolate(method = 'cubic',axis = 0,limit_direction = 'both')
#     counter = counter+1


fig_fairnessNM = go.Figure(data=go.Heatmap(x=x,y=y,z=fairnessNM,colorscale='Jet',zmax = 0.6, zmin = 0))
fig_fairnessNM.update_layout(title='Fairness (%) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_fairnessNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=fairnessNNM,colorscale='Jet',zmax = 0.6, zmin = 0))
fig_fairnessNNM.update_layout(title='Fairness (%) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_taxNM = go.Figure(data=go.Heatmap(x=x,y=y,z=taxNM,colorscale='Jet',zmax = 0.04, zmin = 0))
fig_taxNM.update_layout(title='Tax avoidance (%) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_taxNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=taxNNM,colorscale='Jet',zmax = 0.04, zmin = 0))
fig_taxNNM.update_layout(title='Tax avoidance (%) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_CO2NM = go.Figure(data=go.Heatmap(x=x,y=y,z=carboneNM,colorscale='Jet',zmax = 1200, zmin = 350))
fig_CO2NM.update_layout(title='CO2 avoided (kg) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_CO2NNM = go.Figure(data=go.Heatmap(x=x,y=y,z=carboneNNM,colorscale='Jet',zmax = 1200, zmin = 350))
fig_CO2NNM.update_layout(title='CO2 avoided (kg) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_totNM = go.Figure(data=go.Heatmap(x=x,y=y,z=totNM,colorscale='Jet',zmax = 1800, zmin = 900))
fig_totNM.update_layout(title='Total costs (euros/pers/year) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_totNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=totNNM,colorscale='Jet',zmax = 1800, zmin = 900))
fig_totNNM.update_layout(title='Total costs (euros/pers/year) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_PVNM = go.Figure(data=go.Heatmap(x=x,y=y,z=PVNM,colorscale='Jet',zmax = 9, zmin = 0))
fig_PVNM.update_layout(title='PV capacity (kW) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_PVNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=PVNNM,colorscale='Jet',zmax = 9, zmin = 0))
fig_PVNNM.update_layout(title='PV capacity (kW) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_BatNM = go.Figure(data=go.Heatmap(x=x,y=y,z=BatNM,colorscale='Jet',zmax = 3, zmin = 0))
fig_BatNM.update_layout(title='Battery capacity (kWh) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_BatNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=BatNNM,colorscale='Jet',zmax = 3, zmin = 0))
fig_BatNNM.update_layout(title='Battery capacity (kWh) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_P2HNM = go.Figure(data=go.Heatmap(x=x,y=y,z=P2HNM,colorscale='Jet',zmax = 0.65, zmin = 0))
fig_P2HNM.update_layout(title='Heat-Pump capacity (kW-h) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_P2HNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=P2HNNM,colorscale='Jet',zmax = 0.65, zmin = 0))
fig_P2HNNM.update_layout(title='Heat-Pump capacity (kW-h) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_HSNM = go.Figure(data=go.Heatmap(x=x,y=y,z=HSNM,colorscale='Jet',zmax = 0.35, zmin = 0))
fig_HSNM.update_layout(title='Heat Storage capacity (kWh) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_HSNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=HSNNM,colorscale='Jet',zmax = 0.35, zmin = 0))
fig_HSNNM.update_layout(title='Heat Storage capacity (kWh) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_PLNM = go.Figure(data=go.Heatmap(x=x,y=y,z=PLNM,colorscale='Jet',zmax = 1, zmin = -3))
fig_PLNM.update_layout(title='Relative peak load reduction (%) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_PLNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=PLNNM,colorscale='Jet',zmax = 1, zmin = -3))
fig_PLNNM.update_layout(title='Relative peak load reduction (%) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_PROtotNM = go.Figure(data=go.Heatmap(x=x,y=y,z=PROtotNM,colorscale='Jet',zmax = 1800, zmin = 900))
fig_PROtotNM.update_layout(title='Total costs prosumer (euros/pers/year) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_PROtotNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=PROtotNNM,colorscale='Jet',zmax = 1800, zmin = 900))
fig_PROtotNNM.update_layout(title='Total costs prosumer (euros/pers/year) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_PAStotNM = go.Figure(data=go.Heatmap(x=x,y=y,z=PAStotNM,colorscale='Jet',zmax = 1800, zmin = 900))
fig_PAStotNM.update_layout(title='Total costs passive (euros/pers/year) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_PAStotNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=PAStotNNM,colorscale='Jet',zmax = 1800, zmin = 900))
fig_PAStotNNM.update_layout(title='Total costs passive (euros/pers/year) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_VNPNM = go.Figure(data=go.Heatmap(x=x,y=y,z=VNPNM,colorscale='Jet',zmax = 0.2, zmin = 0))
fig_VNPNM.update_layout(title='VNP (euros/kWh) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_VNPNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=VNPNNM,colorscale='Jet',zmax = 0.2, zmin = 0))
fig_VNPNNM.update_layout(title='VNP (euros/kWh) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_CNPNM = go.Figure(data=go.Heatmap(x=x,y=y,z=CNPNM,colorscale='Jet',zmax = 300, zmin = 0))
fig_CNPNM.update_layout(title='CNP (euros/kW) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_CNPNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=CNPNNM,colorscale='Jet',zmax = 300, zmin = 0))
fig_CNPNNM.update_layout(title='CNP (euros/kW) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_FNPNM = go.Figure(data=go.Heatmap(x=x,y=y,z=FNPNM,colorscale='Jet',zmax = 404, zmin = 0))
fig_FNPNM.update_layout(title='FNP (euros) NM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

fig_FNPNNM = go.Figure(data=go.Heatmap(x=x,y=y,z=FNPNNM,colorscale='Jet',zmax = 404, zmin = 0))
fig_FNPNNM.update_layout(title='FNP (euros) NNM', autosize=False, title_font_size = 25, font_size=22,
                width=900, height=700,
                margin=dict(l=65, r=50, b=65, t=90), 
                yaxis = dict(
                        title = 'Volumetric Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = y_axis,
                        tickvals= y_axis),
                        #ticksuffix=' Hours'),
                xaxis = dict(
                        title = 'Capacity Share (%)',
                        nticks = 5,
                        #tickmode = 'array',
                        ticktext = x_axis,
                        tickvals= x_axis)
                )

# fig_fairnessNM.write_html(Folder + "Figure/"+Add_name+"HM_fairnessNM.html")
# fig_fairnessNNM.write_html(Folder + "Figure/"+Add_name+"HM_fairnessNNM.html")
# fig_taxNM.write_html(Folder + "Figure/"+Add_name+"HM_taxNM.html")
# fig_taxNNM.write_html(Folder + "Figure/"+Add_name+"HM_taxNNM.html")
# fig_CO2NM.write_html(Folder + "Figure/"+Add_name+"HM_CO2NM.html")
# fig_CO2NNM.write_html(Folder + "Figure/"+Add_name+"HM_CO2NNM.html")
# fig_totNM.write_html(Folder + "Figure/"+Add_name+"HM_totNM.html")
# fig_totNNM.write_html(Folder + "Figure/"+Add_name+"HM_totNNM.html")
# fig_PVNM.write_html(Folder + "Figure/"+Add_name+"HM_PVNM.html")
# fig_PVNNM.write_html(Folder + "Figure/"+Add_name+"HM_PVNNM.html")
# fig_BatNM.write_html(Folder + "Figure/"+Add_name+"HM_BatNM.html")
# fig_BatNNM.write_html(Folder + "Figure/"+Add_name+"HM_BatNNM.html")
# fig_P2HNM.write_html(Folder + "Figure/"+Add_name+"HM_P2HNM.html")
# fig_P2HNNM.write_html(Folder + "Figure/"+Add_name+"HM_P2HNNM.html")
# fig_HSNM.write_html(Folder + "Figure/"+Add_name+"HM_HSNM.html")
# fig_HSNNM.write_html(Folder + "Figure/"+Add_name+"HM_HSNNM.html")
# fig_PLNM.write_html(Folder + "Figure/"+Add_name+"HM_PLNM.html")
# fig_PLNNM.write_html(Folder + "Figure/"+Add_name+"HM_PLNNM.html")
# fig_PROtotNM.write_html(Folder + "Figure/"+Add_name+"HM_PROtotNM.html")
# fig_PROtotNNM.write_html(Folder + "Figure/"+Add_name+"HM_PROtotNNM.html")
# fig_PAStotNM.write_html(Folder + "Figure/"+Add_name+"HM_PAStotNM.html")
# fig_PAStotNNM.write_html(Folder + "Figure/"+Add_name+"HM_PAStotNNM.html")
# fig_VNPNM.write_html(Folder + "Figure/"+Add_name+"HM_VNPNM.html")
# fig_VNPNNM.write_html(Folder + "Figure/"+Add_name+"HM_VNPNNM.html")
# fig_CNPNM.write_html(Folder + "Figure/"+Add_name+"HM_CNPNM.html")
# fig_CNPNNM.write_html(Folder + "Figure/"+Add_name+"HM_CNPNNM.html")
# fig_FNPNM.write_html(Folder + "Figure/"+Add_name+"HM_FNPNM.html")
# fig_FNPNNM.write_html(Folder + "Figure/"+Add_name+"HM_FNPNNM.html")

fig_fairnessNM.write_image(Folder + "Figure/"+Add_name+"HM_fairnessNM.pdf")
fig_fairnessNNM.write_image(Folder + "Figure/"+Add_name+"HM_fairnessNNM.pdf")
fig_taxNM.write_image(Folder + "Figure/"+Add_name+"HM_taxNM.pdf")
fig_taxNNM.write_image(Folder + "Figure/"+Add_name+"HM_taxNNM.pdf")
fig_CO2NM.write_image(Folder + "Figure/"+Add_name+"HM_CO2NM.pdf")
fig_CO2NNM.write_image(Folder + "Figure/"+Add_name+"HM_CO2NNM.pdf")
fig_totNM.write_image(Folder + "Figure/"+Add_name+"HM_totNM.pdf")
fig_totNNM.write_image(Folder + "Figure/"+Add_name+"HM_totNNM.pdf")
fig_PVNM.write_image(Folder + "Figure/"+Add_name+"HM_PVNM.pdf")
fig_PVNNM.write_image(Folder + "Figure/"+Add_name+"HM_PVNNM.pdf")
fig_BatNM.write_image(Folder + "Figure/"+Add_name+"HM_BatNM.pdf")
fig_BatNNM.write_image(Folder + "Figure/"+Add_name+"HM_BatNNM.pdf")
fig_P2HNM.write_image(Folder + "Figure/"+Add_name+"HM_P2HNM.pdf")
fig_P2HNNM.write_image(Folder + "Figure/"+Add_name+"HM_P2HNNM.pdf")
fig_HSNM.write_image(Folder + "Figure/"+Add_name+"HM_HSNM.pdf")
fig_HSNNM.write_image(Folder + "Figure/"+Add_name+"HM_HSNNM.pdf")
fig_PLNM.write_image(Folder + "Figure/"+Add_name+"HM_PLNM.pdf")
fig_PLNNM.write_image(Folder + "Figure/"+Add_name+"HM_PLNNM.pdf")
fig_PROtotNM.write_image(Folder + "Figure/"+Add_name+"HM_PROtotNM.pdf")
fig_PROtotNNM.write_image(Folder + "Figure/"+Add_name+"HM_PROtotNNM.pdf")
fig_PAStotNM.write_image(Folder + "Figure/"+Add_name+"HM_PAStotNM.pdf")
fig_PAStotNNM.write_image(Folder + "Figure/"+Add_name+"HM_PAStotNNM.pdf")
fig_VNPNM.write_image(Folder + "Figure/"+Add_name+"HM_VNPNM.pdf")
fig_VNPNNM.write_image(Folder + "Figure/"+Add_name+"HM_VNPNNM.pdf")
fig_CNPNM.write_image(Folder + "Figure/"+Add_name+"HM_CNPNM.pdf")
fig_CNPNNM.write_image(Folder + "Figure/"+Add_name+"HM_CNPNNM.pdf")
fig_FNPNM.write_image(Folder + "Figure/"+Add_name+"HM_FNPNM.pdf")
fig_FNPNNM.write_image(Folder + "Figure/"+Add_name+"HM_FNPNNM.pdf")