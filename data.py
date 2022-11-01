import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl

ElectrolyserPower = 300        # in kW

xls = pd.read_excel('inverter_20221029_204628.xlsx', engine = 'openpyxl')

Power = xls["AC출력(kW)"]

class electrolyzer:
    def __init__(self,production_rate,max_power):
        self.production_rate = production_rate     #kg 
        self.max_power = max_power                 #in kw
        
class ITM_PEM(electrolyzer):
    def __init__(self,production_rate, max_power):
        super().__init__(production_rate, max_power)
        self.min_power = max_power*10/100  # in kW
    
class battery:
    def __init__(self, capacity):                     # kwh
        self.power_rating = 0.25 * capacity           # kwh
        self.capacity = capacity *3600                # kj

electrolyser = ITM_PEM(4.86, ElectrolyserPower)

T = len(Power)                           # no. of hourly periods

state = np.ones((T,1)) 

grid_demand = np.zeros((T,1))            # power required from grid 
H_produced = np.zeros((T,1))             # kg of hydrogen produced in the hour period
batt_power = np.zeros((T,1))             # Amt of power load into battery in kW
batt_soc = np.zeros((T,1))               # Amt of energy stored in battery in kJ
excess_solar = np.zeros((T,1))           # kW
electrolyzer_power = Power

#batt = battery(100)

batt_sol = []
baso = []
gridd = []
for i in range(400):
# Records hydrogen produced in each hour period
    print(i)
    #electrolyser = ITM_PEM(4.86, i)
    batt = battery(i)
    for j in range(T):
        #if electrolyser is on
        if state[j] == 1:
            if electrolyzer_power[j] >= electrolyser.max_power:                              #if theoretical power exceeds max power
                batt_power[j] = min((electrolyzer_power[j] - electrolyser.max_power), batt.power_rating)
                #print(batt_power)
                if batt_power[j] == batt.power_rating:
                    excess_solar[j] = electrolyzer_power[j] - batt.power_rating - electrolyser.max_power
                    #print(excess_solar)
                batt_soc[j] = min((batt_soc[j-1] + batt_power[j]*60*60), batt.capacity)
                
                if batt_soc[j-1] == batt.capacity:
                    excess_solar[j] = electrolyzer_power[j] - electrolyser.max_power
                    #print(excess_solar)
            else:                                                                           #if theoretical power is below min power
                grid_demand[j] = electrolyser.max_power - electrolyzer_power[j] 

                if grid_demand[j] > batt.power_rating:                                 #if power demand exceeds battery power rating
                    if batt_soc[j-1] > batt.power_rating*60*60:
                        batt_soc[j] = batt_soc[j-1] - batt.power_rating*3600
                        grid_demand[j] = grid_demand[j] - batt.power_rating
                    else:
                        batt_soc[j] = batt_soc[j-1]
                elif grid_demand[j] < batt.power_rating:                               #if power demand is within battery power rating
                    #if there is sufficient energy in battery storage
                    if batt_soc[j-1] > grid_demand[j]*60*60:
                        batt_soc[j] = batt_soc[j-1] - grid_demand[j]*3600
                        grid_demand[j] = 0
                    else:
                        batt_soc[j] = batt_soc[j-1]

        #if electrolyser is off
        else:
            H_produced[j] = 0

            if batt_soc[j-1] == batt.capacity:
                excess_solar[j] = electrolyzer_power[j]
                batt_soc[j] = batt_soc[j-1]
            else:
                if electrolyzer_power[j] > batt.power_rating:
                    excess_solar[j] = electrolyzer_power[j] - batt.power_rating
                    batt_soc[j] = min((batt_soc[j-1] + batt.power_rating*60*60), batt.capacity)
                else:
                    batt_soc[j] = min((batt_soc[j-1] + electrolyzer_power[j]*60*60), batt.capacity)
    batt_sol.append(int(np.average(excess_solar)))
    baso.append(int(np.max(batt_soc)))
    gridd.append(int(np.max(grid_demand)))
    #print(batt_soc)
        
plt.figure(figsize = (10,10))

#ax = plt.subplot(1,1,1)
plt.plot(batt_sol)
plt.title("Hourly Battery State of Charge")
plt.xlabel('Batt cap')
plt.ylabel('soc')
plt.savefig("bcs.png")

plt.figure(figsize = (10,10))
plt.plot(batt_sol)
plt.title("Hourly Battery State of Charge")
plt.xlabel('Batt cap')
plt.ylabel('excess sol')
print(type(batt_sol))
print(batt_sol.index(min(batt_sol)))
plt.savefig("bces.png")
print(baso.index(min(gridd)))