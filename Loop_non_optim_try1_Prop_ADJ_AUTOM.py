# -*- coding: utf-8 -*-
"""
Created on Fri Feb 19 10:26:15 2021

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

# ----------------------
# Import of data
# ----------------------

Sets = pd.read_csv('D:/DTU/2A/MSC Thesis/Data/Data_checked/Sets.csv', sep=',',header=0)

# Demand
Electricity_Demand = pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/Av_p4_hNeh_iX_aX_v0.xlsx',header=0)
Heat_Demand = pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/BC_Heat_Demand.xlsx',header=0)

# Energy tariffs
EL_Tariffs= pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/ELprice.xlsx',header=0)
Gas_Tariffs= pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/Gas_price.xlsx',header=0)
DH_Tariffs= pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/DH_price.xlsx',header=0)
Taxes= pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/Taxes.xlsx',header=None, dtype={0: str}).set_index(0).squeeze().to_dict()
CO2_intensities= pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/Intensities.xlsx',index_col=0,header=0,sheet_name='Intensities (kg per kWh)')


# Technology parameters
PV_par=pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/PV_par.xlsx', header=None, dtype={0: str}).set_index(0).squeeze().to_dict()
Battery_par = pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/Battery_par.xlsx', header=None, dtype={0: str}).set_index(0).squeeze().to_dict()  
Grid_par = pd.read_csv('D:/DTU/2A/MSC Thesis/Data/Data_checked/Grid_par.csv', header=None, dtype={0: str}).set_index(0).squeeze().to_dict()  
Boiler_par=pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/Boiler_par.xlsx', header=None, dtype={0: str}).set_index(0).squeeze().to_dict()
P2H_par=pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/P2H_par.xlsx', header=None, dtype={0: str}).set_index(0).squeeze().to_dict()
Heat_Storage_par=pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/Heat_Storage_par.xlsx', header=None, dtype={0: str}).set_index(0).squeeze().to_dict()
DH_par=pd.read_csv('D:/DTU/2A/MSC Thesis/Data/Data_checked/DH_par.csv', header=None, dtype={0: str}).set_index(0).squeeze().to_dict()

# Capacity factors
PV_CF = pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/SolarCF.xlsx',index_col=0,header=0)

# Designs to test
Design_test = pd.read_excel('D:/DTU/2A/MSC Thesis/Data/Data_checked/Design_test.xlsx',header = 0)

# Time
T=[i for i in range(0,len(Sets))]
# T = [i for i in range(0,200)]
 
for bl in range(0,len(Design_test)):
    # ----------------------------------------------------------------------------
    # Preparation of the results and loop
    # ----------------------------------------------------------------------------
    # Consumer data randomly selected
    Nbr_Cons = 2
    Pro_Pen = 0.5
    Consumers = pd.DataFrame(0, index=[i for i in range(Nbr_Cons)], columns = ['Percentage','Heat Demand','Electricity Demand','Active/Passive'])
    for i in range(Nbr_Cons):
        Consumers.loc[i,'Heat Demand'] = 'P' + str(3)
        Consumers.loc[i,'Electricity Demand'] = 'T' + str(1)
        if i == 0:
            Consumers.loc[i,'Active/Passive'] = 1
            Consumers.loc[i,'Percentage'] = Pro_Pen
        else:
            Consumers.loc[i,'Active/Passive'] = 0
            Consumers.loc[i,'Percentage'] = 1-Pro_Pen
      
    # Creation of the tabe storing the results time after time
    Results_NC = pd.DataFrame(0,[i for i in range(Nbr_Cons)], columns = ['Network charges collected','PV','Battery','P2H','HS','Boiler','Total Network Cost','Taxes','O&M','Energy costs','Investments','Carbon cost','Volumetric NC part','Capacity NC part','Fixed NC part','PV prod','G import','G export','DH import','Bo prod','P2H Hprod'])
    NC_collected = 0 # NC collected --> careful should equal charges for the DSO by people
    Sunk_Costs = 404*100
    
    Results_time = pd.DataFrame(0,index = [t for t in range(0,len(T))],columns = [])
    for i in range(0,Nbr_Cons):
        for t in range(0,len(T)):
            Results_time.loc[t,'Import'+str(i)] = 0
            Results_time.loc[t,'Export'+str(i)] = 0
            Results_time.loc[t,'PV prod'+str(i)] = 0
            Results_time.loc[t,'Bat ch'+str(i)] = 0
            Results_time.loc[t,'Bat dh'+str(i)] = 0
            Results_time.loc[t,'Bat st'+str(i)] = 0
            Results_time.loc[t,'P2H ELcons'+str(i)] = 0
            Results_time.loc[t,'DH Import'+str(i)] = 0
            # Results_time.loc[t,'P2H Hprod'+str(i)] = 0
            Results_time.loc[t,'Bo prod'+str(i)] = 0
            # Results_time.loc[t,'HS ch'+str(i)] = 0
            # Results_time.loc[t,'HS dh'+str(i)] = 0
            # Results_time.loc[t,'HS st'+str(i)] = 0
    
    Costs_time = pd.DataFrame(0,index = [t for t in range(0,len(T))],columns = [])
    for i in range(0,Nbr_Cons):
        for t in range(0,len(T)):
            Costs_time.loc[t,'Network Costs'+str(i)] = 0
            # Costs_time.loc[t,'Energy Costs'+str(i)] = 0
    #         Costs_time.loc[t,'O&M Costs'+str(i)] = 0
    #         Costs_time.loc[t,'Taxes'+str(i)] = 0
    #         Costs_time.loc[t,'Carbone Costs'+str(i)] = 0
    
    # creation of the table helping to compare the result and exit the loop
    Results_NC_prev = pd.DataFrame(0,[i for i in range(Nbr_Cons)], columns = ['Network charges collected','PV','Battery','P2H','HS','Boiler'])
    Results_NC_prev.loc[i,'Network charges collected'] = 0
    Results_NC_prev.loc[i,'PV']=100
    Results_NC_prev.loc[i,'Battery']=100
    Results_NC_prev.loc[i,'P2H']=100
    Results_NC_prev.loc[i,'HS']=100
    Results_NC_prev.loc[i,'Boiler']=100
    
    # Initial values input in the program
    Param_prev = pd.DataFrame(0,index=['Values'], columns=['NM','NNM','VNP','CNP','FNP','Vol_per','Capa_per','min_adjust','max_adjust'])
    Param_prev.loc['Values','NM'] = Design_test.loc[bl,'NM'] # Net Metering
    Param_prev.loc['Values','NNM'] = Design_test.loc[bl,'NNM'] # Not Net Metering
    Param_prev.loc['Values','VNP'] = 10 #Variable Network price in €/kWh
    Param_prev.loc['Values','CNP'] = 10 #Capacity Network Price in €/kW
    Param_prev.loc['Values','FNP'] = 10 #Fixed Network Price
    Param_prev.loc['Values','Vol_per'] = Design_test.loc[bl,'Vol_per'] #percentage of volumetric charges in network tariffs
    Param_prev.loc['Values','Capa_per'] = Design_test.loc[bl,'Capa_per'] #percentage of capacity charges in network tariffs
    BC = 1
    min_adjustment = 0.05
    max_adjustment = 3
    
    profile_combined = pd.DataFrame(0,index=[i for i in range(0,len(T))],columns = ['Profile combined'])
    for i in range(0,Nbr_Cons):
        for t in range(0,len(T)):
            profile_combined.loc[t,'Profile combined'] = profile_combined.loc[t,'Profile combined'] + Electricity_Demand.loc[t,Consumers.loc[i,'Electricity Demand']]
    
    quantile = profile_combined.quantile(q=[0.95,0.05])
    
    Price_Corr = pd.DataFrame(0,index=[i for i in range(0,len(T))],columns = ['Price Correction'])
    for t in range(0,len(T)):
        if profile_combined.loc[t,'Profile combined'] <= quantile.loc[0.05,'Profile combined']:
            Price_Corr.loc[t,'Price Correction'] = min_adjustment
        elif profile_combined.loc[t,'Profile combined'] >= quantile.loc[0.95,'Profile combined']:
            Price_Corr.loc[t,'Price Correction'] = max_adjustment
        else:
            Price_Corr.loc[t,'Price Correction'] = 1
    
    Carbon_price = 0.02 # €/kg
    Gas_Intensity = 0.201959999983843 #kg/kWh
    Sets_tech = ['PV','Battery','P2H','HS','Boiler']
    
    
    Relative_change_capa = 1
    it = 0
    it_mem = pd.DataFrame(0,index=[it],columns=['Iteration','Relative Capacity Change'])
    # ----------------------------------------------------------------------------
    # Loop
    # ----------------------------------------------------------------------------
    while abs(Relative_change_capa) > 0.0001:    
        for i in range(Nbr_Cons):
            Demand_Type = Consumers.loc[i,'Electricity Demand']
            Heat_Type = Consumers.loc[i,'Heat Demand']
            AorP = Consumers.loc[i,'Active/Passive']
            # ----------------------
            # Building of the model
            # ----------------------
            
            M = ConcreteModel()
            
            # ----------------------
            # Declaration of parameters and sets 
            # ----------------------
            
            # Time parameter, one year, hour step
            M.T=Set(doc='hour of the year',initialize=T)
            
            # Calculation tool
            M.flag_1=Param(initialize=1,doc='Flag for energy fluxes (1)')
            M.flag_0=Param(initialize=0,doc='Flag for energy fluxes (0)')
            NM = Param_prev.loc['Values','NM']
            NNM = Param_prev.loc['Values','NNM']
            VNP = Param_prev.loc['Values','VNP']
            CNP = Param_prev.loc['Values','CNP']
            FNP = Param_prev.loc['Values','FNP']
            Vol_per = Param_prev.loc['Values','Vol_per']
            Capa_per= Param_prev.loc['Values','Capa_per']       
            # ----------------------
            # Declaration of variables and parameters linked to technologies
            # ----------------------
            
            # Grid parameters and variables
            M.G_flag=Var(M.T,domain=Binary,doc='Import/Export Flag')
            M.G_export=Var(M.T,doc='Export to the grid in kW',initialize=0,bounds=(0,float(Grid_par["Ex_lim"])))
            M.G_import=Var(M.T,doc='Import from the grid in kW',initialize=0,bounds=(0,float(Grid_par["Im_lim"])))
            
            if AorP == 1:
                # PV parameters and variables
                M.PV_Capa=Var(initialize=20,bounds=(0,100),doc='Capacity of installed PV kW')
                M.PV_prod=Var(M.T,domain=NonNegativeReals,initialize=0)
                
                # # Battery Parameters and variables
                M.B_Capa=Var(initialize=20,bounds=(0,100),doc='Capacity of installed battery kWh')
                M.B_charge=Var(M.T,domain=NonNegativeReals,doc='Battery charging in hour T',initialize=0)
                M.B_discharge=Var(M.T,domain=NonNegativeReals,doc='Battery discharging in hour T',initialize=0)
                M.B_status=Var(M.T,domain=NonNegativeReals,doc='Battery status in hour T',initialize=0)
                
                # # P2H parameters
                M.P2H_Capa=Var(initialize=20,bounds=(0,100),doc='Capacity of installed P2H kW')
                M.P2H_Hprod=Var(M.T,domain=NonNegativeReals,initialize=0)
                M.P2H_ELcons=Var(M.T,domain=NonNegativeReals,initialize=0)
                
            if AorP == 0:
                # PV parameters and variables
                M.PV_Capa=Param(initialize=0,doc='Capacity of installed PV kW')
                M.PV_prod=Var(M.T,domain=NonNegativeReals,initialize=0)
                
                # # Battery Parameters and variables
                M.B_Capa=Param(initialize=0,doc='Capacity of installed battery kWh')
                M.B_charge=Var(M.T,domain=NonNegativeReals,doc='Battery charging in hour T',initialize=0)
                M.B_discharge=Var(M.T,domain=NonNegativeReals,doc='Battery discharging in hour T',initialize=0)
                M.B_status=Var(M.T,domain=NonNegativeReals,doc='Battery status in hour T',initialize=0)
                
                # # P2H parameters
                M.P2H_Capa=Param(initialize=0,doc='Capacity of installed P2H kW')
                M.P2H_Hprod=Var(M.T,domain=NonNegativeReals,initialize=0)
                M.P2H_ELcons=Var(M.T,domain=NonNegativeReals,initialize=0) 
            
            # # Boilers parameters
            M.Bo_Capa=Var(initialize=20,bounds=(0,100),doc='Capacity of installed gas boiler kW')
            M.Bo_prod=Var(M.T,domain=NonNegativeReals,initialize=0)
            
            # # Heat Storage parameters
            M.HS_Capa=Var(initialize=0,bounds=(0,100),doc='Capacity of installed Heat Storage kWh')
            M.HS_charge=Var(M.T,domain=NonNegativeReals,doc='Heat Storage charging in hour T',initialize=0)
            M.HS_discharge=Var(M.T,domain=NonNegativeReals,doc='Heat Storage discharging in hour T',initialize=0)
            M.HS_status=Var(M.T,domain=NonNegativeReals,doc='Heat Storage status in hour T',initialize=0)
            
            # # District Heating parameters and variables
            M.DH_import=Var(M.T,doc='Import from the DH in kW',initialize=0,bounds=(0,float(DH_par["Im_lim"])))
            
            M.Peak_Load=Var(doc='Peak Load',initialize=0)
            # ----------------------
            # Objective function
            # ----------------------
            
            # Definition of the objective of the model
            def Carbone_emitted_Cost(M):
                return sum(M.G_import[t]*float(CO2_intensities.loc[t,'EL average'])*Carbon_price + M.DH_import[t]*float(CO2_intensities.loc[t,'H average'])*Carbon_price + M.Bo_prod[t]*Gas_Intensity*Carbon_price for t in T)
                # return 0
            
            def Energy_Cost_c(M):
                # Energy cost for one consumer 
                return sum((M.G_import[t] - M.G_export[t])*float(EL_Tariffs.loc[t,"Import"]) + M.DH_import[t]*float(DH_Tariffs.loc[t,"Import"]) + M.Bo_prod[t]*float(Gas_Tariffs.loc[0,'Value']) for t in T) 
            
            def Network_Charges_c(M):
                # Network charges for one consumer 
                if AorP == 1:
                    return sum((M.G_import[t]-M.G_export[t]*(NM-NNM))*VNP*Price_Corr.loc[t,'Price Correction'] for t in T) + M.Peak_Load*CNP + FNP #add eventually the proportions
                else:
                    return sum((M.G_import[t]-M.G_export[t]*(NM-NNM))*VNP for t in T) + M.Peak_Load*CNP + FNP
                # return 0
            
            def Fixed_Charges_c(M):
                #Charges that are the same for everyone
                return 0
            
            def Investment_Cost_c(M):
                # Investment cost in the equipment for one consumer
                # Annuity factor ? 
                PV_CAPEX_annualized = float(PV_par["Capital_cost"])/float(PV_par["Lifetime"])
                Battery_CAPEX_annualized = float(Battery_par["Capital_cost"])/float(Battery_par["Lifetime"])
                P2H_CAPEX_annualized = float(P2H_par["Capital_cost"])/float(P2H_par["Lifetime"])
                HS_CAPEX_annualized = float(Heat_Storage_par["Capital_cost"])/float(Heat_Storage_par["Lifetime"])
                Boiler_CAPEX_annualized = float(Boiler_par["Capital_cost"])/float(Boiler_par["Lifetime"])
                return PV_CAPEX_annualized*M.PV_Capa + Battery_CAPEX_annualized*M.B_Capa + P2H_CAPEX_annualized*M.P2H_Capa + Boiler_CAPEX_annualized*M.Bo_Capa + HS_CAPEX_annualized*M.HS_Capa
                # return 
            
            def OM_cost_c(M):
                return sum(M.PV_prod[t]*float(PV_par['OM_cost']) + (M.B_charge[t]+M.B_discharge[t])*float(Battery_par['OM_cost']) + M.Bo_prod[t]*float(Boiler_par['OM_cost']) + (M.HS_charge[t]+M.HS_discharge[t])*float(Heat_Storage_par['OM_cost']) + M.P2H_Hprod[t]*float(P2H_par['OM_cost'])  for t in T) 
                # return 0
            
            def Tax_cost_c(M):
                return sum((M.G_import[t] + M.PV_prod[t] - M.G_export[t] - M.P2H_ELcons[t])*float(Taxes['Electricity']) + M.P2H_ELcons[t]*float(Taxes['Electricity_for_heat']) + M.DH_import[t]*float(Taxes['Heat']) + M.Bo_prod[t]*float(Taxes['Gas']) for t in T)
                # return 0
            
            # Total costs for the consumer
            # def Objective_rule(M):
            #     return Energy_Cost_c(M) + Network_Charges_c(M) + Fixed_Charges_c(M) + Investment_Cost_c(M) + OM_cost_c(M) + Tax_cost_c(M) + Carbone_emitted_Cost(M)
            # M.Objective_Rule=Objective(rule=Objective_rule,sense=minimize)
            Objective_rule = Energy_Cost_c(M) + Network_Charges_c(M) + Fixed_Charges_c(M) + Investment_Cost_c(M) + OM_cost_c(M) + Tax_cost_c(M) + Carbone_emitted_Cost(M)
            M.Objective_Rule=Objective(expr=Objective_rule,sense=minimize)
            
            
            # ----------------------
            # Constraints
            # ----------------------
            
            # PV constraints -------------------------------------------------------------
            def PV_production_rule(M,t):
                # Estimate the production of power through PV in kW
                return M.PV_prod[t] == M.PV_Capa*float(PV_CF.loc[t,'SolarCF'])
            M.PV_prod_Cons=Constraint(M.T,rule=PV_production_rule)
            # ----------------------------------------------------------------------------
            
            # P2H constraints -------------------------------------------------------------
            def P2H_production_rule(M,t):
                return M.P2H_ELcons[t] == M.P2H_Hprod[t]/P2H_par['Eff']
            M.P2H_prod_Cons=Constraint(M.T,rule=P2H_production_rule)
            
            def P2H_production_max_rule(M,t):
                return M.P2H_Hprod[t] <= M.P2H_Capa
            M.P2H_prod_max_Cons=Constraint(M.T,rule=P2H_production_max_rule)
            # ----------------------------------------------------------------------------
            
            
            # Energy Demand constraints -------------------------------------------------
            def EBalance_Demand_rule(M,t):
                #make sure that grid, battery and production combined can meet the demand
                return float(Electricity_Demand.loc[t,Demand_Type]) == M.G_import[t] - M.G_export[t] + M.B_discharge[t] - M.B_charge[t] + M.PV_prod[t] - M.P2H_ELcons[t]
            M.EBalance_Demand_Cons=Constraint(M.T,rule=EBalance_Demand_rule)
            
            def HBalance_Demand_rule(M,t):
                #make sure that boiler and heat storage combined can meet the demand
                return float(Heat_Demand.loc[t,Heat_Type]) == M.DH_import[t] + M.Bo_prod[t] - M.HS_charge[t] + M.HS_discharge[t] + M.P2H_Hprod[t]
            M.HBalance_Demand_Cons=Constraint(M.T,rule=HBalance_Demand_rule)
            # ----------------------------------------------------------------------------
            
            # Boiler constraints -----------------------------------------------------------
            def Bo_production_max_rule(M,t):
                return M.Bo_prod[t] <= M.Bo_Capa
            M.Bo_prod_max_Cons=Constraint(M.T,rule=Bo_production_max_rule)
            # ----------------------------------------------------------------------------
            
            # # Battery constraints --------------------------------------------------------
            def B_max_discharge(M,t):
                # Makes sure that cannot discharge more than what is in the battery the previous time
                if t == 0:
                    return M.B_discharge[t] <= M.B_Capa*(1-float(Battery_par["Linear_Losses"])) 
                else:
                    return M.B_discharge[t] <= M.B_status[t-1]*(1-float(Battery_par["Linear_Losses"])) 
            M.B_max_discharge_Cons=Constraint(M.T,rule=B_max_discharge)
            
            def battery_max_state(M,t):
                #Battery is charged in the maximum capacity
                return M.B_status[t] <= M.B_Capa*float(Battery_par["Max_charge"])
            M.B_Status_Max_Cons=Constraint(M.T,rule=battery_max_state)
            
            def battery_min_state(M,t):
                #Battery is charged over the minimum capacity
                return M.B_status[t] >= (1-float(Battery_par["Depth_of_discharge"]))*M.B_Capa
            M.B_Status_Min_Cons=Constraint(M.T,rule=battery_min_state)
            
            def BFollowing_charge(M,t):
                # Makes sure that the losses are taken into account inside the battery
                if t == 0:
                    return M.B_status[t] == M.B_Capa*(1-float(Battery_par["Linear_Losses"])) - M.B_discharge[t]/float(Battery_par["Discharging_eff"]) + M.B_charge[t]*float(Battery_par["Charging_eff"])
                else:
                    return M.B_status[t] == M.B_status[t-1]*(1-float(Battery_par["Linear_Losses"])) - M.B_discharge[t]/float(Battery_par["Discharging_eff"]) + M.B_charge[t]*float(Battery_par["Charging_eff"])
            M.B_status_Cons=Constraint(M.T,rule=BFollowing_charge)
            
            def B_max_discharge2(M,t):
                # Makes sure that cannot discharge more than what is in the battery the previous time
                return M.B_discharge[t] <= M.B_Capa * float(Battery_par["Discharging_lim"])
            M.B_max_discharge2_Cons=Constraint(M.T,rule=B_max_discharge2)
            
            def B_max_charge2(M,t):
                # Makes sure that cannot discharge more than what is in the battery the previous time
                return M.B_charge[t] <= M.B_Capa * float(Battery_par["Charging_lim"])
            M.B_max_charge2_Cons=Constraint(M.T,rule=B_max_charge2)
            # ----------------------------------------------------------------------------
            
            # Heat Storage constraints --------------------------------------------------------
            def HS_max_discharge(M,t):
                # Makes sure that cannot discharge more than what is in the battery the previous time
                if t == 0:
                    return M.HS_discharge[t] <= M.HS_Capa*(1-float(Heat_Storage_par["Linear_Losses"])) 
                else:
                    return M.HS_discharge[t] <= M.HS_status[t-1]*(1-float(Heat_Storage_par["Linear_Losses"])) 
            M.HS_max_discharge_Cons=Constraint(M.T,rule=HS_max_discharge)
            
            def Heat_Storage_max_state(M,t):
                #Heat Storage is charged in the maximum capacity
                return M.HS_status[t] <= M.HS_Capa*float(Heat_Storage_par["Max_charge"])
            M.HS_Status_Max_Cons=Constraint(M.T,rule=Heat_Storage_max_state)
            
            def Heat_Storage_min_state(M,t):
                #Battery is charged over the minimum capacity
                return M.HS_status[t] >= (1-float(Heat_Storage_par["Depth_of_discharge"]))*M.HS_Capa
            M.HS_Status_Min_Cons=Constraint(M.T,rule=Heat_Storage_min_state)
            
            def HSFollowing_charge(M,t):
                # Makes sure that the losses are taken into account inside the battery
                if t == 0:
                    return M.HS_status[t] == M.HS_Capa*(1-float(Heat_Storage_par["Linear_Losses"])) - M.HS_discharge[t]/float(Heat_Storage_par["Discharging_eff"]) + M.HS_charge[t]*float(Heat_Storage_par["Charging_eff"])
                else:
                    return M.HS_status[t] == M.HS_status[t-1]*(1-float(Heat_Storage_par["Linear_Losses"])) - M.HS_discharge[t]/float(Heat_Storage_par["Discharging_eff"]) + M.HS_charge[t]*float(Heat_Storage_par["Charging_eff"])
            M.HS_status_Cons=Constraint(M.T,rule=HSFollowing_charge)
            
            def HS_max_discharge2(M,t):
                # Makes sure that cannot discharge more than what is in the battery the previous time
                return M.HS_discharge[t] <= M.HS_Capa * float(Heat_Storage_par["Discharging_lim"])
            M.HS_max_discharge2_Cons=Constraint(M.T,rule=HS_max_discharge2)
            
            def HS_max_charge2(M,t):
                # Makes sure that cannot discharge more than what is in the battery the previous time
                return M.HS_charge[t] <= M.HS_Capa * float(Heat_Storage_par["Charging_lim"])
            M.HS_max_charge2_Cons=Constraint(M.T,rule=HS_max_charge2)
            # ----------------------------------------------------------------------------
            
            # Grid constraints -----------------------------------------------------------
            
            # ----------------------------------------------------------------------------
            
            # DH constraints -----------------------------------------------------------
            
            # ----------------------------------------------------------------------------
            
            #General constraints ----------------------------------------------------------
            def Peak_Load_Rule(M,t):
                return M.G_import[t] + M.G_export[t] <= M.Peak_Load
            M.Peak_Load_cons=Constraint(M.T,rule=Peak_Load_Rule)
            
            def Network_Charges_rule(M):
                # Network charges for one consumer 
                if AorP == 1:
                    return sum((M.G_import[t] - M.G_export[t])*Price_Corr.loc[t,'Price Correction'] for t in T) >= 0
                else:
                    return sum(M.G_import[t] - M.G_export[t] for t in T) >= 0
            M.Network_Charges_cons=Constraint(rule=Network_Charges_rule)
            # -----------------------------------------------------------------------------
            
            # ----------------------
            # Solver
            # ----------------------
            
            solver = SolverFactory("gurobi", solver_io="python")
            # solver=SolverFactory('glpk')
            Res_Obj=solver.solve(M)
            print(value(M.Objective_Rule))
            Results_NC.loc[i,'Network charges collected'] = value(Network_Charges_c(M))
            Results_NC.loc[i,'PV']=value(M.PV_Capa)
            Results_NC.loc[i,'Battery']=value(M.B_Capa)
            Results_NC.loc[i,'P2H']=value(M.P2H_Capa)
            Results_NC.loc[i,'HS']=value(M.HS_Capa)
            Results_NC.loc[i,'Boiler']=value(M.Bo_Capa)
            Results_NC.loc[i,'Peak Load'] = value(M.Peak_Load)
            Results_NC.loc[i,'Total Network Cost'] = value(M.Objective_Rule)
            Results_NC.loc[i,'Taxes'] = value(Tax_cost_c(M))
            Results_NC.loc[i,'O&M'] = value(OM_cost_c(M))
            Results_NC.loc[i,'Energy costs'] = value(Energy_Cost_c(M))
            Results_NC.loc[i,'Investments'] = value(Investment_Cost_c(M))
            Results_NC.loc[i,'Carbon cost'] = value(Carbone_emitted_Cost(M))
            Results_NC.loc[i,'PV prod'] = value(sum(M.PV_prod[t] for t in T))
            Results_NC.loc[i,'G import'] = value(sum(M.G_import[t] for t in T))
            Results_NC.loc[i,'G export'] = value(sum(M.G_export[t] for t in T))
            Results_NC.loc[i,'DH import'] = value(sum(M.DH_import[t] for t in T))
            Results_NC.loc[i,'Bo prod'] = value(sum(M.Bo_prod[t] for t in T))
            Results_NC.loc[i,'P2H Hprod'] = value(sum(M.P2H_Hprod[t] for t in T))
            
            
            for t in range(0,len(T)):
                Results_time.loc[t,'Import'+str(i)] = value(M.G_import[t])
                Results_time.loc[t,'Export'+str(i)] = value(M.G_export[t])
                Results_time.loc[t,'PV prod'+str(i)] = value(M.PV_prod[t])
                Results_time.loc[t,'Bat ch'+str(i)] = value(M.B_charge[t])
                Results_time.loc[t,'Bat dh'+str(i)] = value(M.B_discharge[t])
                Results_time.loc[t,'Bat st'+str(i)] = value(M.B_status[t])
                Results_time.loc[t,'P2H ELcons'+str(i)] = value(M.P2H_ELcons[t])
                Results_time.loc[t,'DH Import'+str(i)] = value(M.DH_import[t])
                # Results_time.loc[t,'P2H Hprod'+str(i)] = value(M.P2H_Hprod[t])
                Results_time.loc[t,'Bo prod'+str(i)] = value(M.Bo_prod[t])
                # Results_time.loc[t,'HS ch'+str(i)] = value(M.HS_charge[t])
                # Results_time.loc[t,'HS dh'+str(i)] = value(M.HS_discharge[t])
                # Results_time.loc[t,'HS st'+str(i)] = value(M.HS_status[t])
                # Costs_time.loc[t,'Energy Costs'+str(i)] = value((M.G_import[t] - M.G_export[t])*float(EL_Tariffs.loc[t,"Import"]) + M.DH_import[t]*float(DH_Tariffs.loc[t,"Import"]) + M.Bo_prod[t]*float(Gas_Tariffs.loc[0,'Value']))
                # Costs_time.loc[t,'O&M Costs'+str(i)] = value(M.PV_prod[t]*float(PV_par['OM_cost']) + (M.B_charge[t]+M.B_discharge[t])*float(Battery_par['OM_cost']) + M.Bo_prod[t]*float(Boiler_par['OM_cost']) + (M.HS_charge[t]+M.HS_discharge[t])*float(Heat_Storage_par['OM_cost']) + M.P2H_Hprod[t]*float(P2H_par['OM_cost']))
                # Costs_time.loc[t,'Taxes'+str(i)] = value((M.G_import[t] + M.PV_prod[t] - M.G_export[t] - M.P2H_ELcons[t])*float(Taxes['Electricity']) + M.P2H_ELcons[t]*float(Taxes['Electricity_for_heat']) + M.DH_import[t]*float(Taxes['Heat']) + M.Bo_prod[t]*float(Taxes['Gas']))
                # Costs_time.loc[t,'Carbone Costs'+str(i)] = value(M.G_import[t]*float(CO2_intensities.loc[t,'EL average'])*Carbon_price + M.DH_import[t]*float(CO2_intensities.loc[t,'H average'])*Carbon_price)
                
    
        # Net metering or not net metering
        DSO_NM = float(Param_prev.loc['Values','NM'])
        DSO_NNM = float(Param_prev.loc['Values','NNM'])
        DSO_VNP = float(Param_prev.loc['Values','VNP'])
        DSO_CNP = float(Param_prev.loc['Values','CNP'])
        DSO_FNP = float(Param_prev.loc['Values','FNP'])
        DSO_Vol_per = float(Param_prev.loc['Values','Vol_per'])
        DSO_Capa_per= float(Param_prev.loc['Values','Capa_per'])
        DSO_Sunk_Costs = float(Sunk_Costs)
        
        # ------------------------------------------------------------------------
        # Objective function for the DSO
        # ------------------------------------------------------------------------
        def VCosts_collected_rule():
            VCosts = 0
            for i in range(Nbr_Cons):
                AorP = Consumers.loc[i,'Active/Passive']
                if AorP == 1:
                    VCosts = VCosts + sum((float(Results_time.loc[t,'Import'+str(i)]) - float(Results_time.loc[t,'Export'+str(i)])*(DSO_NM - DSO_NNM))*Price_Corr.loc[t,'Price Correction'] for t in T)*Consumers.loc[i,'Percentage']*100
                else:
                    VCosts = VCosts + sum((float(Results_time.loc[t,'Import'+str(i)]) - float(Results_time.loc[t,'Export'+str(i)])*(DSO_NM - DSO_NNM)) for t in T)*Consumers.loc[i,'Percentage']*100
            return VCosts
            # return sum(sum((float(Results_time.loc[t,'Import'+str(i)]) - float(Results_time.loc[t,'Export'+str(i)])*(DSO_NM - DSO_NNM)) for t in T)*Consumers.loc[i,'Percentage']*100 for i in range(0,Nbr_Cons)) #*float(Consumers.loc[i,'Percentage'])*DSO_Vol_per
    
        
        def CCosts_collected_rule():
            return sum(float(Results_NC.loc[i,'Peak Load'])*Consumers.loc[i,'Percentage']*100 for i in range(0,Nbr_Cons)) 
        
        def FCosts_collected_rule():
            return sum(Consumers.loc[i,'Percentage']*100 for i in range(0,Nbr_Cons)) 
        
        
        # ------------------------------------------------------------------------
        # Constraints
        # ------------------------------------------------------------------------
        DSO_VNP = DSO_Sunk_Costs * DSO_Vol_per / VCosts_collected_rule()
        DSO_CNP = DSO_Sunk_Costs * DSO_Capa_per / CCosts_collected_rule()
        DSO_FNP = DSO_Sunk_Costs * (1 - DSO_Vol_per - DSO_Capa_per) / FCosts_collected_rule()
        
        
        Param_prev.loc['Values','NM'] = DSO_NM
        Param_prev.loc['Values','NNM'] = DSO_NNM
        Param_prev.loc['Values','VNP'] = DSO_VNP
        Param_prev.loc['Values','CNP'] = DSO_CNP
        Param_prev.loc['Values','FNP'] = DSO_FNP
        Param_prev.loc['Values','Vol_per'] = DSO_Vol_per
        Param_prev.loc['Values','Capa_per'] = DSO_Capa_per
        
        if sum(sum(Results_NC_prev.loc[cons,tech] for tech in Sets_tech) for cons in range(0,Nbr_Cons)) == 0:
            Relative_change_capa = sum(sum((Results_NC.loc[cons,tech]-Results_NC_prev.loc[cons,tech]) for tech in Sets_tech) for cons in range(0,Nbr_Cons))
        else: 
            Relative_change_capa = sum(sum((Results_NC.loc[cons,tech]-Results_NC_prev.loc[cons,tech]) for tech in Sets_tech) for cons in range(0,Nbr_Cons)) / sum(sum(Results_NC_prev.loc[cons,tech] for tech in Sets_tech) for cons in range(0,Nbr_Cons))
        
        it = it + 1
        it_mem.loc[it-1,'Iteration'] = it
        it_mem.loc[it-1,'Relative Capacity Change'] = Relative_change_capa
        
        for i in range(0,Nbr_Cons):
            AorP = Consumers.loc[i,'Active/Passive']
            Results_NC_prev.loc[i,'Network charges collected'] = Results_NC.loc[i,'Network charges collected']
            Results_NC_prev.loc[i,'PV']=Results_NC.loc[i,'PV']
            Results_NC_prev.loc[i,'Battery']=Results_NC.loc[i,'Battery']
            Results_NC_prev.loc[i,'P2H']=Results_NC.loc[i,'P2H']
            Results_NC_prev.loc[i,'HS']=Results_NC.loc[i,'HS']
            Results_NC_prev.loc[i,'Boiler']=Results_NC.loc[i,'Boiler']
            if AorP ==1:
                Results_NC.loc[i,'Volumetric NC part'] = sum((float(Results_time.loc[t,'Import'+str(i)]) - float(Results_time.loc[t,'Export'+str(i)])*(DSO_NM - DSO_NNM))*DSO_VNP*Price_Corr.loc[t,'Price Correction'] for t in T)
            if AorP == 0:
                Results_NC.loc[i,'Volumetric NC part'] = sum((float(Results_time.loc[t,'Import'+str(i)]) - float(Results_time.loc[t,'Export'+str(i)])*(DSO_NM - DSO_NNM))*DSO_VNP for t in T)
            Results_NC.loc[i,'Capacity NC part'] = float(Results_NC.loc[i,'Peak Load'])*DSO_CNP
            Results_NC.loc[i,'Fixed NC part'] = DSO_FNP
            Results_NC.loc[i,'Network Costs Recovered'] = (Results_NC.loc[i,'Fixed NC part'] + Results_NC.loc[i,'Capacity NC part'] + Results_NC.loc[i,'Volumetric NC part'])*Consumers.loc[i,'Percentage']*100
            for t in range(0,len(T)):
                if AorP == 1:
                    Costs_time.loc[t,'Network Costs'+str(i)] = (float(Results_time.loc[t,'Import'+str(i)]) - float(Results_time.loc[t,'Export'+str(i)])*(DSO_NM - DSO_NNM))*DSO_VNP*Price_Corr.loc[t,'Price Correction'] + (float(Results_NC.loc[i,'Peak Load'])*DSO_CNP + DSO_FNP)/len(T)
                if AorP == 0:
                    Costs_time.loc[t,'Network Costs'+str(i)] = (float(Results_time.loc[t,'Import'+str(i)]) - float(Results_time.loc[t,'Export'+str(i)])*(DSO_NM - DSO_NNM))*DSO_VNP + (float(Results_NC.loc[i,'Peak Load'])*DSO_CNP + DSO_FNP)/len(T)
    
        
    # ----------------------------------------------------------------------------
    # Result processing
    # ----------------------------------------------------------------------------
    Outcome_res = pd.DataFrame(0,index=[i for i in range(0,Nbr_Cons)], columns = ['CO2 avoided (kg)'])
    # CO2 avoided emissions
    for i in range(0,Nbr_Cons):
        Demand_Type = Consumers.loc[i,'Electricity Demand']
        Heat_Type = Consumers.loc[i,'Heat Demand']
        AorP = Consumers.loc[i,'Active/Passive']
        Outcome_res.loc[i,'CO2 avoided (kg)'] = sum((float(Electricity_Demand.loc[t,Demand_Type])-Results_time.loc[t,'Import'+str(i)])*float(CO2_intensities.loc[t,'EL average']) + (float(Heat_Demand.loc[t,Heat_Type])-Results_time.loc[t,'DH Import'+str(i)])*float(CO2_intensities.loc[t,'H average']) - Results_time.loc[t,'Bo prod'+str(i)]*Gas_Intensity for t in T)
            
               
    Param_prev.loc['Values','min_adjust'] = min_adjustment
    Param_prev.loc['Values','max_adjust'] = max_adjustment    
    
    # # ----------------------
    # # Export of results
    # # ----------------------
    
    name = 'D:/DTU/2A/MSC Thesis/Model/New/'+ 'ADJ' + 'BC' + str(BC) + '_al' + str(Param_prev.loc['Values','Vol_per']) + '_b' + str(Param_prev.loc['Values','Capa_per']) + '_NM' + str(Param_prev.loc['Values','NM']) + '_Nbr' + str(Nbr_Cons) + '_ProProp' + str(Pro_Pen) + '_SC' + str(Sunk_Costs) + '_CO2C' + str(Carbon_price) + 'min_adj' + str(min_adjustment) + 'max_adj' + str(max_adjustment) + '.xlsx'
    with pd.ExcelWriter(name) as writer:
        Consumers.to_excel(writer,sheet_name='Consumer details')
        Results_NC.to_excel(writer, sheet_name='Capacities and costs')
        Results_time.to_excel(writer, sheet_name='Time values')
        Param_prev.to_excel(writer, sheet_name='DSO parameters')
        Outcome_res.to_excel(writer,sheet_name = 'Outcome of the optimization')
        it_mem.to_excel(writer,sheet_name = 'Convergence')
        # Costs_time.to_excel(writer,sheet_name = 'Costs timely')
    # excel_writer = 'D:\DTU\2A\MSC Thesis\Model\New'