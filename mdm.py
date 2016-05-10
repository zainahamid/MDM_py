import pyexcel
from pyexcel.ext import xlsx
from pyexcel.ext import xls
import os
import scipy.optimize as opt 
import numpy as np

#Take in global background content from csv files - fe & strategy
with open(os.path.join('fe.csv'), 'r') as f:
        FE_data = f.readlines()
    
#with open(os.path.join('specs.csv'), 'r') as f:
#        Specs_data = f.readlines()
        
for i in range(len(FE_data)):
    FE_data[i] = FE_data[i].replace('\n','').replace('\r','').split(',')
#for i in range(len(Specs_data)):
#    Specs_data[i] = Specs_data[i].replace('\n','').replace('\r','').split(',')

#book now contains the entire excel workbook 
dict = pyexcel.get_book_dict(file_name="WG_input.xls")
macro = np.asarray(dict['macro'])
vehdata = np.asarray(dict['VEHDATA'])
lgroup = np.asarray(dict['LGROUP'])
tgroup = np.asarray(dict['tGROUP'])
fgroup = np.asarray(dict['fGROUP'])
sgroup = np.asarray(dict['sGROUP'])
oegroup = np.asarray(dict['oeGROUP'])
bgroup = np.asarray(dict['bGROUP'])
scenario = np.asarray(dict['scenario'])
constraints = np.asarray(dict['Constraints'])
Basedelta = np.asarray(dict['Basedelta'])

book = pyexcel.get_book(file_name="WG_input.xls")
startYear = 2015

temp = book['macro']
temp.name_columns_by_row(0)
macro_rec = temp.to_records()

temp = book['VEHDATA']
temp.name_columns_by_row(0)
vehdata_rec = temp.to_records()

temp = book['LGROUP']
temp.name_columns_by_row(0)
lgroup_rec = temp.to_records()

temp = book['tGROUP']
temp.name_columns_by_row(0)
tgroup_rec = temp.to_records()

temp = book['fGROUP']
temp.name_columns_by_row(0)
fgroup_rec = temp.to_records()

temp = book['sGROUP']
temp.name_columns_by_row(0)
sgroup_rec = temp.to_records()

temp = book['oeGROUP']
temp.name_columns_by_row(0)
oegroup_rec = temp.to_records()

temp = book['bGROUP']
temp.name_columns_by_row(0)
bgroup_rec = temp.to_records()

row_size = len(vehdata)
col_size = len(vehdata[0])
targetVehData_rec = []
list_vehData = []
list_vehData.append(vehdata[0])

#calculating the target for the year we are working on
for i in range(len(vehdata_rec)):
    if int(vehdata_rec[i]['year']) == startYear:
        targetVehData_rec.append(vehdata_rec[i])
        list_vehData.append(vehdata[i+1])
        print 'done'
    
targetVehData = np.asarray(list_vehData)