import yaml
import os, fnmatch
from collections import defaultdict 
from openpyxl import load_workbook
import pandas as pd 
from openpyxl.utils.dataframe import dataframe_to_rows

# Update This
toolkit_file = "resources/SAMM_spreadsheet.xlsx"
result_file = "result/final.xlsx"
data_files_dir = "resources/model"
yml_file_pattern = '*.yml'


# Load SAMM
practice_level_files = fnmatch.filter(os.listdir(data_files_dir+"/practice_levels"), yml_file_pattern)
question_files = fnmatch.filter(os.listdir(data_files_dir+"/questions"), yml_file_pattern)
stream_files = fnmatch.filter(os.listdir(data_files_dir+"/streams"), yml_file_pattern)
activity_files = fnmatch.filter(os.listdir(data_files_dir+"/activities"), yml_file_pattern)
practice_files = fnmatch.filter(os.listdir(data_files_dir+"/security_practices"), yml_file_pattern)
function_files = fnmatch.filter(os.listdir(data_files_dir+"/business_functions"), yml_file_pattern)
answer_set_files = fnmatch.filter(os.listdir(data_files_dir+"/answer_sets"), yml_file_pattern)
maturity_level_files = fnmatch.filter(os.listdir(data_files_dir+"/maturity_levels"), yml_file_pattern)

def loadDict(dataFiles, paths, splitnum):
    temp = defaultdict(list)
    
    for z in paths:
        with open(dataFiles+'/'+z) as file:
            temp[z.split(' ')[splitnum].split('.')[0]] = yaml.full_load(file)

    return temp

def getID(myDict, find_key, value, return_key):
    for z in myDict.keys():
        if find_key in myDict[z]:
            if myDict[z][find_key] == value:
                if return_key != '':
                    return myDict[z][return_key]
                else:
                    return z

def getIndex(myArray,value):
    for idx,z in enumerate(myArray):
        if value in z:
            return idx

# Define all elements
# pL = defaultdict(list) 
# q = defaultdict(list) 
# s = defaultdict(list) 
# a = defaultdict(list) 
# p = defaultdict(list) 
# f = defaultdict(list) 
# aS = defaultdict(list) 
# m = defaultdict(list) 

# # Load YAML files
# pL = loadDict(data_files_dir+"/practice_levels", practice_level_files, 0) 
# q = loadDict(data_files_dir+"/questions", question_files, 1) 
# s = loadDict(data_files_dir+"/streams", stream_files, 1) 
# a = loadDict(data_files_dir+"/activities", activity_files, 1) 
# p = loadDict(data_files_dir+"/security_practices", practice_files, 1) 
# f = loadDict(data_files_dir+"/business_functions", function_files, 1) 
# aS = loadDict(data_files_dir+"/answer_sets", answer_set_files, 1) 
# m = loadDict(data_files_dir+"/maturity_levels", maturity_level_files, 1) 

# # Create Answer Sets DataFrame

# asColumns = ['ANS_SET_CODE','A','B','C','D','A_W','B_W','C_W','D_W']

# answer_set_files = []
# asData = []

# for i in sorted (aS.keys()):
#     temp = []
#     asKeys = [ 'text','value']
#     temp.append(i)
#     for y in asKeys:
#         for x in range(4):
#             temp.append(aS[i]['values'][x][y])
#     answer_set_files.append(temp)
    

# for idx,z in enumerate(answer_set_files):
#     asData.append([idx,"No"] + z[2:])


# asDF = pd.DataFrame(asData,columns = asColumns)

# # Create Questions DataFrame

# qColumns = ['ID','Business Function','Security Practice','Activity','Maturity','Question','Guidance','Answer Option']
# qData = []

# question_files = []
# for t in sorted (q.keys()):
#     temp=[]
#     temp.append(t.split('-')[0]+'-'+t.split('-')[1]+'-'+t.split('-')[3]+'-'+t.split('-')[2])
#     temp.append(getID(f,'id',getID(p,'id',getID(s,'id',getID(a,'id',q[t]['activity'],'stream'), 'practice'),'function'),'name'))
#     temp.append(getID(p,'id',getID(s,'id',getID(a,'id',q[t]['activity'],'stream'), 'practice'),'name'))
#     temp.append(getID(s,'id',getID(a,'id',q[t]['activity'],'stream'), 'name'))
#     temp.append(getID(m,'id',getID(pL,'id',getID(a,'id',q[t]['activity'],'level'),'maturitylevel'),'number'))
#     temp.append(q[t]['text'])
#     temp.append(q[t]['quality'])
#     temp.append(getIndex(answer_set_files,getID(aS,'id',q[t]['answerset'],'')))
#     question_files.append(temp)
    
# for z in question_files:
#     tmp = []
#     tmp.append(z[0]+'-1')
#     tmp.append(z[1])
#     tmp.append(z[2])
#     tmp.append(z[3])
#     tmp.append(z[4])
#     tmp.append(z[5])
#     tmp.append("\015 ".join(z[6]))
#     tmp.append(z[7])

#     qData.append(tmp)

# qDF = pd.DataFrame(qData,columns = qColumns)

# Open Worksheet
wb = load_workbook(toolkit_file)
# wbquestion = wb['imp-questions']
# wbanswers = wb['imp-answers']

# Clear out old data (may need to update ranges if SAMM changes dramatically)
# wbquestion.delete_cols(1,20)
# wbquestion.delete_rows(1,250)
# wbanswers.delete_cols(1,20)
# wbanswers.delete_rows(1,250)

# Drop DataFrames into the Toolkit
# for q in dataframe_to_rows(qDF, index=False, header=True):
#     wbquestion.append(q)

# for a in dataframe_to_rows(asDF, index=False, header=True):
#     wbanswers.append(a)

# Save!!!
wb.save(result_file)