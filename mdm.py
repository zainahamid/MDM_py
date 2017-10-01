import pyexcel
import scipy.optimize as opt 
import numpy as np
import xlsxwriter

cafe_average =  {'domestic':60, 'Asian': 100, 'European':70}

#book now contains the entire excel workbook 
dict = pyexcel.get_book_dict(file_name="Input_Data.xls")

tgroup = np.asarray(dict['tGROUP'])
fgroup = np.asarray(dict['fGROUP'])
bgroup = np.asarray(dict['bGROUP'])
scenario = np.asarray(dict['Scenario'])

book = pyexcel.get_book(file_name="Input_Data.xls")
startYear = 2015
endYear = 2017
num_of_iterations = 50
gamma = 0.1

#order by row 0
temp = book['Scenario_2']
temp.name_columns_by_row(0)
scenario_rec = temp.to_records()

def calculate_shares(scenario_rec):
    #s0 = sum of each sj
    s0 = 0.0
    for eachrec in scenario_rec:
        s0+=eachrec['sj']
    s0 = 1 - s0
    
    #DELTA CALCULATION FOLLOWED BY SHARE CALCULATION
    sums={}
    for eachrec in scenario_rec:
        if 'delta' not in eachrec.keys():
            delta = 0.0
            
        delta = eachrec['phi'] * np.log(eachrec['sj']) + (eachrec['rho'] - eachrec['phi']) * np.log(eachrec['sfu']) + (eachrec['sigma'] - eachrec['rho']) * np.log(eachrec['st']) + (1 - eachrec['sigma']) * np.log(eachrec['sb']) - np.log(s0) - eachrec['alpha'] * eachrec['price']
        eachrec['delta'] = delta
        
        #needs to be done every iteration
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
        #print (eachrec['modname'], eachrec['new_shares'], eachrec['new_shares']*eachrec['population'])
        #print('\n')
        
#create a list of unique OEM names
oems = []
shares_oems={}
calculate_shares(scenario_rec)
for eachrec in scenario_rec:
    if eachrec['oem'] not in oems:
        oems.append(eachrec['oem'])
    if eachrec['oem'] not in shares_oems.keys():
        shares_oems[eachrec['oem']] = copy.deepcopy(scenario_rec)
        
workbook = xlsxwriter.Workbook('results.xlsx')
for year in range (startYear, endYear+1):
    gamma = 0.1
    test_prices={}
    
    print('%%%$$$%%%$$$%%%$$$')
    print(year)
    print('%%%$$$%%%$$$%%%$$$')
    print('\n')
        
    for i in range (num_of_iterations):
        gamma = gamma * 0.95     
        #calculate_shares(scenario_rec)#the global record of scn
        #update the global & the share
        for each_oem in oems:
            #create a temp state of shares
            shares_oems[each_oem] = copy.deepcopy(scenario_rec)
            share_oem = shares_oems[each_oem]
            
        #OPTIMIZATION
        for each_oem in oems:            
            flag = 0
            bnds = ()
            A, C = [], []
            arow = []
            sum_q, bnumerator, bdenom = 0, 0, 0
            for eachrec in share_oem:
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
            A = []
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
                print ('Average CAFE value: ', avg / sum_q)
                print (each_oem, res['x'], C, A, B, bnds)
                print ('\n')
                
                print ('OEM: ',each_oem)
                matched_rec, main_rec = 0, 0
                for eachrec in share_oem:
                    if eachrec['oem'] == each_oem:
                        if eachrec['modname'] not in test_prices.keys():
                            test_prices[eachrec['modname']] = list()
                        test_prices[eachrec['modname']].append({'prices':eachrec['price_2'],'demand':eachrec['new_shares']*eachrec['population'], 'optimal':res['x'][matched_rec], 'bounds': bnds[matched_rec]})
                        factor = -1 * int(np.log10(abs(res['x'][matched_rec] - eachrec['new_shares']*eachrec['population'])))
                        factor_2 = gamma * 10**factor * (eachrec['new_shares']*eachrec['population'] - res['x'][matched_rec])
                        print(eachrec['modname'])
                        print('Demand',eachrec['new_shares']*eachrec['population'])
                        print('Optimal',res['x'][matched_rec])
                        print('bounds',bnds[matched_rec] )
                        print ('gamma', gamma)
                        print('diff', (res['x'][matched_rec] - eachrec['new_shares']*eachrec['population']))
                        print ('FAAAAA', factor)
                        print ('factor', factor_2)
                        print('Old Price:', eachrec['price_2']) 
                        #Check over here if this new demand value & optimal value difference is less OR if demand > upper bound, continue with the prev demand, prices
                        if len(test_prices[eachrec['modname']]) > 1:
                            if eachrec['new_shares']*eachrec['population'] > bnds[matched_rec][1] or eachrec['new_shares']*eachrec['population'] < bnds[matched_rec][0] or abs(eachrec['new_shares']*eachrec['population'] -  res['x'][matched_rec]) > abs(test_prices[eachrec['modname']][-2]['demand'] - test_prices[eachrec['modname']][-2]['optimal']):
                                eachrec['price_2'] = test_prices[eachrec['modname']][-2]['prices']
                                eachrec['new_shares'] = test_prices[eachrec['modname']][-2]['demand']/eachrec['population']
                                #make current prices same as old prices
                                #make current demand same as old demand 
                            else:
                                eachrec['price_2']+=factor_2         
                        else:
                            eachrec['price_2']+=factor_2 
                        print('New Price:', eachrec['price_2'])
                        print('\n')
                        scenario_rec[main_rec]['price_2'] = eachrec['price_2']
                        matched_rec+=1
                    main_rec+=1
                        
            #done going over that oem for that iteration
            calculate_shares(scenario_rec)
                        
                        
    #Update all the values for the next year's calculation
                        
    worksheet = workbook.add_worksheet('Year_'+str(year))
    row, col = 0, 0

    sum_sfu, sum_st, sum_sb = {}, {}, {}
    for eachrec in scenario_rec:
        for record in sorted(eachrec.keys()):
            worksheet.write(row, col, record)
            col+=1
        row+=1
        break
    
    for eachrec in scenario_rec:
        col = 0
        eachrec['volume'] = eachrec['new_shares']*eachrec['population']
        eachrec['price'] = eachrec['price_2']
        eachrec['sj'] = eachrec['new_shares']
        eachrec['ProdMin']=eachrec['volume'] * 0.7
        eachrec['ProdMax']=eachrec['volume'] * 1.3
        for record in sorted(eachrec.keys()):
            worksheet.write(row, col, eachrec[record])
            col+=1
        row+=1
        
        if eachrec['fuGroupId'] not in sum_sfu.keys():
            sum_sfu[eachrec['fuGroupId']] = eachrec['sj']
        else:
            sum_sfu[eachrec['fuGroupId']] += eachrec['sj']
            
        if eachrec['tGroupId'] not in sum_st.keys():
            sum_st[eachrec['tGroupId']] = eachrec['sj']
        else:
            sum_st[eachrec['tGroupId']] += eachrec['sj']
            
        if eachrec['bGroupId'] not in sum_sb.keys():
            sum_sb[eachrec['bGroupId']] = eachrec['sj']
        else:
            sum_sb[eachrec['bGroupId']] += eachrec['sj']
        
    for eachrec in scenario_rec:
        eachrec['sfu'] = sum_sfu[eachrec['fuGroupId']]
        eachrec['st'] =  sum_st[eachrec['tGroupId']]
        eachrec['sb'] = sum_sb[eachrec['bGroupId']]
workbook.close()
    

        
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
        
