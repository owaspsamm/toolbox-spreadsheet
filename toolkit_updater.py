import yaml
import os, fnmatch
from collections import defaultdict 
from openpyxl import load_workbook
import pandas as pd 
from openpyxl.utils.dataframe import dataframe_to_rows

# Update This
toolkit_file = "/opt/app/resources/SAMM_spreadsheet.xlsx"
result_file = "/github/workspace/SAMM_spreadsheet.xlsx"
data_files_dir = "/github/workspace/model"


def load_dictionary(path):
    yaml_files = fnmatch.filter(os.listdir(path), '*.yml')
    temp = defaultdict(list)
    
    for yaml_file in yaml_files:
        with open(path+"/"+yaml_file) as file:
            temp[yaml_file.split('.')[0]] = yaml.full_load(file)
    return temp

def get_id(my_dictionary, find_key, value, return_key):
    for key in my_dictionary.keys():
        if find_key in my_dictionary[key] and my_dictionary[key][find_key] == value:
            if return_key != '':
                return my_dictionary[key][return_key]
            else:
                return key

def get_index(my_array,value):
    for idx,z in enumerate(my_array):
        if value in z:
            return idx


# # Load YAML files
pL = load_dictionary(data_files_dir+"/practice_levels") 
q = load_dictionary(data_files_dir+"/questions") 
s = load_dictionary(data_files_dir+"/streams") 
a = load_dictionary(data_files_dir+"/activities") 
p = load_dictionary(data_files_dir+"/security_practices") 
f = load_dictionary(data_files_dir+"/business_functions") 
aS = load_dictionary(data_files_dir+"/answer_sets") 
m = load_dictionary(data_files_dir+"/maturity_levels") 

# Create Answer Sets DataFrame

as_columns = ['ANS_SET_CODE','A','B','C','D','A_W','B_W','C_W','D_W']

answer_set_files = []
as_data = []

for i in sorted (aS.keys()):
    temp = []
    asKeys = [ 'text','value']
    temp.append(i)
    for y in asKeys:
        for x in range(4):
            temp.append(aS[i]['values'][x][y])
    answer_set_files.append(temp)
    

for idx,z in enumerate(answer_set_files):
    as_data.append([idx,"No"] + z[2:])

as_df = pd.DataFrame(as_data,columns = as_columns)

# Create Questions DataFrame

q_columns = ['ID','Business Function','Security Practice','Activity','Maturity','Question','Guidance','Answer Option']
q_data = []

question_files = []
for t in sorted (q.keys()):
    temp=[]
    temp.append(t.split('-')[0]+'-'+t.split('-')[1]+'-'+t.split('-')[3]+'-'+t.split('-')[2])
    temp.append(get_id(f,'id',get_id(p,'id',get_id(s,'id',get_id(a,'id',q[t]['activity'],'stream'), 'practice'),'function'),'name'))
    temp.append(get_id(p,'id',get_id(s,'id',get_id(a,'id',q[t]['activity'],'stream'), 'practice'),'name'))
    temp.append(get_id(s,'id',get_id(a,'id',q[t]['activity'],'stream'), 'name'))
    temp.append(get_id(m,'id',get_id(pL,'id',get_id(a,'id',q[t]['activity'],'level'),'maturitylevel'),'number'))
    temp.append(q[t]['text'])
    temp.append(q[t]['quality'])
    temp.append(get_index(answer_set_files,get_id(aS,'id',q[t]['answerset'],'')))
    question_files.append(temp)
    
for z in question_files:
    tmp = []
    tmp.append(z[0]+'-1')
    tmp.append(z[1])
    tmp.append(z[2])
    tmp.append(z[3])
    tmp.append(z[4])
    tmp.append(z[5])
    tmp.append("\015 ".join(z[6]))
    tmp.append(z[7])

    q_data.append(tmp)

q_df = pd.DataFrame(q_data,columns = q_columns)
# Open Worksheet
wb = load_workbook(toolkit_file)
wb_questions = wb['imp-questions']
wb_answers = wb['imp-answers']

# Clear out old data (may need to update ranges if SAMM changes dramatically)
wb_questions.delete_cols(1,20)
wb_questions.delete_rows(1,250)
wb_answers.delete_cols(1,20)
wb_answers.delete_rows(1,250)

# Drop DataFrames into the Toolkit
for q in dataframe_to_rows(q_df, index=False, header=True):
    wb_questions.append(q)

for a in dataframe_to_rows(as_df, index=False, header=True):
    wb_answers.append(a)

# Save!!!
wb.save(result_file)
