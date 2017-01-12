import pyexcel
from pyexcel.ext import xlsx
from pyexcel.ext import xls
import os
import scipy.optimize as opt 
import numpy as np
import math
cafe_average =  {'domestic':30, 'Asian': 100, 'European':70}
#book now contains the entire excel workbook 
dict = pyexcel.get_book_dict(file_name="Input_Data.xls")
macro = np.asarray(dict['Macro'])
vehdata = np.asarray(dict['VEHDATA'])
tgroup = np.asarray(dict['tGROUP'])
fgroup = np.asarray(dict['fGROUP'])
bgroup = np.asarray(dict['bGROUP'])
scenario = np.asarray(dict['Scenario'])
constraints = np.asarray(dict['Constraints'])
Basedelta = np.asarray(dict['BASEDELTA'])

book = pyexcel.get_book(file_name="Input_Data.xls")
startYear = 2015
normalFactor = 3000.0 * 1000.0
numVolumeVar = 1
num_of_iterations = 1
gamma = 0.1
#omiga = 0.93

#order by column 0                                             
temp = book['Macro']
temp.name_rows_by_column(0)
macro_rec = temp.to_records()

#order by row 0
temp = book['VEHDATA']
temp.name_columns_by_row(0)
vehdata_rec = temp.to_records()

#order by row 0
temp = book['Scenario_2']
temp.name_columns_by_row(0)
scenario_rec = temp.to_records()

#order by row 0
temp = book['Constraints']
temp.name_columns_by_row(0)
constraints_rec = temp.to_records()

#order by row 0
temp = book['BASEDELTA']
temp.name_columns_by_row(0)
Basedelta_rec = temp.to_records()

#Vehicle Data - "Price" to be read as times 10000
#Constraints - "Cost" to read as times 10000

#create a list of unique OEM names
oems=[]
s0 = 0.0
for eachrec in scenario_rec:
    s0+=eachrec['sj']
s0 = 1 - s0
sums = {}

for eachrec in scenario_rec:
    delta = 0.0
    if eachrec['oem'] not in oems:
        oems.append(eachrec['oem'])
        
    delta = eachrec['phi'] * np.log(eachrec['sj']) + (eachrec['rho'] - eachrec['phi']) * np.log(eachrec['sfu']) + (eachrec['sigma'] - eachrec['rho']) * np.log(eachrec['st']) + (1 - eachrec['sigma']) * np.log(eachrec['sb']) - np.log(s0) - eachrec['alpha'] * eachrec['price']
    #another formula to solve delta
    #delta = np.log(eachrec['sj']) - np.log(s0)- eachrec['alpha'] * eachrec['price'] - (1 - eachrec['phi']) * np.log(eachrec['sj']/eachrec['sfu']) - (1 - eachrec['rho']) * np.log(eachrec['sfu']/eachrec['sb']) - (1 - eachrec['sigma']) * np.log(eachrec['sb']/eachrec['st'])
        
    eachrec['delta'] = delta
    eachrec['vj'] = delta + eachrec['alpha'] * eachrec['price_2']
    
    #Here calculate everything thats needed to calculate share
    eachrec['e_vj_phi'] = np.exp(eachrec['vj']/eachrec['phi'])
    key = 'sum_basegroup' + str(eachrec['fuGroupId'])
    if key not in sums.keys():
        sums[key] = 0.0
    sums[key] += eachrec['e_vj_phi']
   
#after the end of this loop sum_base is calculated for all the 5 base groups
    
#another loop for each value in the sum of base_groups over phi & rho and then their sum
visited = {}
for eachrec in scenario_rec:
    key2 = 'sum_subgroup' + str(eachrec['tGroupId'])
    if key2 not in sums.keys():
        sums[key2] = 0.0
    if 'sum_basegroup'+str(eachrec['fuGroupId']) not in visited.keys():
        sums[key2] += np.power(sums['sum_basegroup'+str(eachrec['fuGroupId'])],(eachrec['phi']/eachrec['rho']))
        visited['sum_basegroup'+str(eachrec['fuGroupId'])] = True
        
#another loop for each value in the sum of sub groups over rho & sigma and their sums
visited = {}
for eachrec in scenario_rec:
    key2 = 'sum_group' + str(eachrec['bGroupId'])
    if key2 not in sums.keys():
        sums[key2] = 0.0
    if 'sum_subgroup'+str(eachrec['tGroupId']) not in visited.keys():
        sums[key2] += np.power(sums['sum_subgroup'+str(eachrec['tGroupId'])],(eachrec['rho']/eachrec['sigma']))
        visited['sum_subgroup'+str(eachrec['tGroupId'])] = True


#share calculation to be verified    
for eachrec in scenario_rec:
    eachrec['new_shares'] = (eachrec['e_vj_phi'] * np.power(sums['sum_basegroup'+str(eachrec['fuGroupId'])],(eachrec['phi']/eachrec['rho']-1)) * np.power(sums['sum_subgroup'+str(eachrec['tGroupId'])],(eachrec['rho']/eachrec['sigma']-1)) * np.power(sums['sum_group'+str(eachrec['bGroupId'])],(eachrec['sigma']-1))) / (np.power(sums['sum_group'+str(eachrec['bGroupId'])],eachrec['sigma'])+1)
    print (eachrec['modname'], eachrec['new_shares'], eachrec['new_shares']*eachrec['population'])
    print('\n')
## MARKET CLEARANCE PART
'''
Repear stuff from above with changed parameters?
Maybe do this n times and reduce the gap between predicted and what clears the
market?
#'''
for i in range (num_of_iterations):
    for each_oem in oems:
        
        flag = 0
        bnds = ()
        A, C = [], []
        arow = []
        sum_q, bnumerator, bdenom = 0, 0, 0
        for eachrec in scenario_rec:
            if eachrec['oem'] == each_oem:
                #C has to have all the profit values that are to MAX for each quantity
                C.append((eachrec['price_2']-eachrec['Cost'])*-1)
                
                #bounds are the respective bounds of quantity
                bound = (int(eachrec['ProdMin']),int(eachrec['ProdMax']))
                bnds = bnds + (bound,)
                arow.append(eachrec['fueleff'])
                #bnumerator += eachrec['new_shares']*eachrec['population'] * eachrec['fueleff']
                #bdenom +=eachrec['fueleff']
                sum_q += eachrec['new_shares']*eachrec['population']
        A=[]
        A.append(arow)
        B = []
        B.append(cafe_average[each_oem] * sum_q)
        
        #A's first row has to have the coefficients of sigma qi = 1         
            #B has the value of (sigma(qi * ei)/sigma(qi))
                
        #maximization
        res = opt.linprog(C, A_ub=A, b_ub=B, bounds = bnds, options={"disp": True})
        avg = 0
        if res['success']:
            for each_res in range(len(res['x'])):
                avg += A[0][each_res] * res['x'][each_res]
            print (avg / sum_q)
        print (each_oem, res['x'], C, A, B, bnds)
        print('\n')
        
#        
#        s0+=eachrec['sj']
#        #if #comparision between sj & sj_new on an average for all is less than 1% break
#        if math.abs(eachrec['sj_new']-eachrec['sj']) / eachrec['sj'] > 0.01:
#            flag = 1
#     
#    if not flag:
#        break
#    s0 = 1 - s0 #calculating a new s0
#    
#    for eachrec in scenario_rec:
#        #replace all sjs with sj_new and calculate the new sj
#        eachrec['sj'] = eachrec['sj_new']
#        
#        delta = np.log(eachrec['sj']) - np.log(s0)
#        - eachrec['alpha'] * eachrec['price']
#        - (1 - eachrec['phi']) * np.log(eachrec['sj']/eachrec['sfu'])
#        - (1 - eachrec['rho']) * np.log(eachrec['sfu']/eachrec['sb'])
#        - (1 - eachrec['sigma']) * np.log(eachrec['sb']/eachrec['st'])
#        
#        eachrec['delta'] = delta
#        eachrec['vj'] = delta + eachrec['alpha'] * eachrec['price']
        
