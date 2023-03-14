# %%
import json
import numpy as np
import pandas as pd
import lxml
import re
import os
from datetime import datetime

# %% [markdown]
# # Script summary
# - Import several Excel files from D365 that lists sales orders with Load IDs and consolidate them in one dataframe
# - Ask the user how many sessions this should be split into
# - Split the dataframe into n x csv files
# - Convert the CSV files to Excel

# %% [markdown]
# ### Clean up the xml_prep folder

# %%
path = "data/rel_to_wh/outbound_to_EA/"
dir_list = os.listdir(path)

for i in range(0,len(dir_list)):
    xml_file = path+dir_list[i]
    os.remove(xml_file)

# %% [markdown]
# ### Import files from D365 containing sales order lines with Load IDs

# %%
path = "data/rel_to_wh/inbound_from_D365/"
dir_list = os.listdir(path)

# %%
print("Reading load files from D365\n")

# %%
loadfiles_concat = pd.read_excel(path+dir_list[0])
input_file = path+dir_list[0]
print(input_file)
try:
    loadfiles_concat.rename(columns={'Load ID': 'LoadID'}, inplace=True)
except:
    pass

# %%
for i in range(1,len(dir_list)):
    input_file = path+dir_list[i]
    print(input_file)
    temp = pd.read_excel(path+dir_list[i])
    try:
        temp.rename(columns={'Load ID': 'LoadID'}, inplace=True)
    except:
        pass
    loadfiles_concat = pd.concat([loadfiles_concat, temp], ignore_index=True)

print('')

# %%
#loadfiles_concat.to_excel('data/rel_to_wh/consolidated.xlsx')

# %%
# loads_to_wh = loadfiles_concat[['LoadID','Description']].copy()
# loads_to_wh.drop(columns={'Description'}, inplace=True, axis=1)

loads_to_wh = loadfiles_concat[['LoadID']].copy()

# %%
loads_to_wh.drop_duplicates(keep='first',inplace=True)
loads_to_wh = loads_to_wh.dropna()
loads_to_wh.reset_index(drop=True, inplace=True)

# %% [markdown]
# ### Split the data among the number of user sessions

# %%
files_str = input('Enter the number of files into which the loads must be split: ')
files = int(files_str)

# %%
for file in range(1,files+1):
    with open('data/rel_to_wh/outbound_to_EA/wh' + str(file) + '.csv', 'a') as fw:        
        fw.write('LoadID'+'\n')

file = 0
for i in range(len(loads_to_wh)):
    if file <= files-1:
        file = file + 1
    else:
        file = 1
    with open('data/rel_to_wh/outbound_to_EA/wh' + str(file) + '.csv', 'a') as fw:        
        fw.write(str(loads_to_wh.loc[i, "LoadID"])+'\n')
        #print(file,loads_to_wh.loc[i, "LoadID"])

# %% [markdown]
# ### Now convert the CSV files into Excel

# %%
path = "data/rel_to_wh/outbound_to_EA/"
dir_list = os.listdir(path)

# %%
# 21 user - Release to warehouse , picking and despatch (Roadnet loads).xlsx

for i in range(0,len(dir_list)):
    excel_file = path+str(i+1)+' user - Release to warehouse , picking and despatch (Roadnet loads).xlsx'
    #print(str(i)+xml_file)
    temp = pd.read_csv(path+dir_list[i])
    temp.to_excel(excel_file, index=False)
    os.remove(path+dir_list[i])

# %%
print('\nOutput files were written to data/rel_to_wh/outbound_to_EA/')


