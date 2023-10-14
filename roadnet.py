# %%
import json
import numpy as np
import pandas as pd
import lxml
import re
import os
from datetime import datetime
import math

# %% [markdown]
# # Script summary
# - Import many xml files that were generated by the Send to Roadnet process in D365
# - Perform processing on the XML files, and import them into a Pandas dataframe containing outbound order lines from D365
# - Convert this into a dataframe that emulates the data that Roadnet would have returned - apply spec for every field in the file
# - Import a list of SessionIDs, one for each user (or virtual user) that will be importing and posting an inbound file from Roadnet
# - Split the inbound file (from Roadnet) into one file per SessionID
# - Export as CSV (take note lines must end with CR-NL)

# %% [markdown]
# ### Clean up temporary folders

# %%
# Set filenames
consolidated_roadnet_out_file = 'data/roadnet/xml_consolidated/consolidated_roadnet_out.xml'

# %%
try:
    os.remove(consolidated_roadnet_out_file)
except:
    pass

# %%
# Clean up the xml_prep folder
path = "data/roadnet/xml_prep/"
dir_list = os.listdir(path)

for i in range(0,len(dir_list)):
    xml_file = path+dir_list[i]
    os.remove(xml_file)

# Clean up the xml_consolidated folder
path = "data/roadnet/xml_consolidated/"
dir_list = os.listdir(path)

for i in range(0,len(dir_list)):
    xml_file = path+dir_list[i]
    os.remove(xml_file)    

# %% [markdown]
# ### Convert D365 XML outbound files from a single line (with no newlines) to multiple lines by inserting a newline after evert '>' character
# - Store the resulting files in the /xml_prep/ folder

# %%
# Get the list of all files and directories
path = "data/roadnet/xml_outbound/"
dir_list = os.listdir(path)

print("Loading " + str(len(dir_list)) + " XML files in ", path)


# %%
def xml_to_multiple_lines(fname, fnum):

    with open(fname, 'r') as fr:
        lines = fr.readlines()
        last_line = len(lines)

    for line in lines:
        replaced_line = re.sub(">", ">\u000A", line)

    outfile = 'data/roadnet/xml_prep/f'+str(fnum)+'.xml'

    with open(outfile, 'w') as fw:
        fw.write(replaced_line)    

    return

# %%
# Convert xml from D365 by adding newlines

for i in range(0,len(dir_list)):
    xml_file = path+dir_list[i]
    xml_to_multiple_lines(xml_file,i)

number_of_xml = len(dir_list)

# %% [markdown]
# ### Consolidate the processed XML outbound files into a single, consolidated xml file
# - Remove all lines that are not transaction line items

# %%
# Get the list of all files and directories
path = "data/roadnet/xml_prep/"
prep_dir_list = os.listdir(path)

# %%
def import_roadnet_files2(fname, fnum, outfile):

    try:    
        with open(fname, 'r') as fr:
            lines = fr.readlines()

            last_line = len(lines)

            with open(outfile, 'a') as fw:
                for line in lines:      
                    substr1 = 'CCBROADNETWORKBENCHSESSIONTABLEENTITY'       
                    x1 = re.search(substr1, line)
                    substr2 = 'Document>'       
                    x2 = re.search(substr2, line)
                    substr3 = 'xml version='       
                    x3 = re.search(substr3, line)
                    if x1 == None and x2 == None and x3 == None:
                        fw.write(line)

    except:
        print("Error importing "+fname)

# %%
for i in range(0,len(prep_dir_list)):
    xml_file = path+prep_dir_list[i]
    import_roadnet_files2(xml_file,i, consolidated_roadnet_out_file)

# %% [markdown]
# ### Split up the consolidated xml file into 5,000 lines, else they cannot be imported into a Pandas dataframe
# - Store these files in the same folder as the consolidated xml, and delete the consolidated xml after the split
# - Add the lines to turn this into a valid XML format

# %%
fname = consolidated_roadnet_out_file
outfile = fname

sizelimit = 5000

try:    
    with open(fname, 'r') as fr:
        lines = fr.readlines()

        last_line = len(lines)

        line_counter = 1

        last_x = 0
        for i in range(0,len(lines)):
            line = lines[i]
            if i == last_line:
                print(line)
            x = int(i/sizelimit)
            with open(outfile[:-4]+str(x)+'.xml', 'a') as fw:        
                if i == 0:
                    fw.write('<?xml version="1.0" encoding="utf-8"?>\n')
                    fw.write('<Document>\n')
                if x > last_x:
                    last_x = x
                    fw.write('<?xml version="1.0" encoding="utf-8"?>\n')
                    fw.write('<Document>\n')
                fw.write(line)
                if int((i+1)/sizelimit) > x:
                    fw.write('</Document>')
                if i == len(lines) - 1:
                    fw.write('</Document>')
    os.remove(consolidated_roadnet_out_file)
                
except:
    print("Error importing "+fname)

# %% [markdown]
# ### Import the transformed XML files into a Pandas dataframe

# %%
# Get the list of all files and directories
path = "data/roadnet/xml_consolidated/"
dir_list = os.listdir(path)

#print("Files and directories in '", path, "' :")

# prints all files
#print(dir_list)

# %%
# Import the first file into a dataframe
rdnet_out = pd.read_xml(path+dir_list[0])

# %%
# Import the rest of the files, and append to the dataframe
for i in range(1,len(dir_list)):
    xml_file = path+dir_list[i]
    print(str(i)+xml_file)
    temp = pd.read_xml(path+dir_list[i])
    rdnet_out = pd.concat([rdnet_out, temp], ignore_index=True)

# %%
rdnet_out.drop(columns={'OUTPERFORMROADNETDESTINATION'}, inplace=True, axis=1)

# %%
rdnet_out = rdnet_out.drop_duplicates(keep='first').copy()

# %% [markdown]
# ### Create the Roadnet inbound file by copying selected columns as-is from the outbound data

# %%
rdnet_in = rdnet_out[['QUANTITY','LOCATIONID','INVENTTRANSID','ITEMID','ORDERID','WAREHOUSEID','PRODUCTNAME','ROADNETROUTE','ORDERACCOUNT','ORDERACCOUNTNAME','WEIGHT']].copy()

# %%
rdnet_in = rdnet_in.dropna(subset=['WAREHOUSEID']).copy()

# %%
rdnet_in.rename(columns={'QUANTITY':'CASEQTY','LOCATIONID':'DESTINATIONLOCATIONID','ORDERID':'ORDERNUMBER','WAREHOUSEID':'ORIGINLOCATIONID','ORDERACCOUNT':'STOPLOCATIONID','ORDERACCOUNTNAME':'STOPLOCATIONNAME'}, inplace=True)

# %% [markdown]
# ### Create the rest of the fields as per the Roadnet inbound file spec

# %%
today = str(datetime.now())
today = today.replace(':','h')
today = today.replace('-','')
today = today.replace(' ','-')
today = today[0:14] + '-'

# %%
rdnet_in['ROADNETROUTEINTERNALROUTEID'] = today + rdnet_in['STOPLOCATIONID'].astype(str)

# %%
no_of_customers = len(rdnet_in['STOPLOCATIONID'].unique())

# %%
# Get legal entity from the invettransid
le_code = rdnet_in.loc[0, 'INVENTTRANSID'][:3]

# %%
# Defaults applied to all LEs

rdnet_in['APPTID'] = ''
rdnet_in['ERROR'] = ''
rdnet_in['LASTSTOPISDESTINATION'] = 'No'
rdnet_in['LOADID'] = ''
rdnet_in['LOADTEMPLATEID'] = ''
rdnet_in['ORDERTYPE'] = 'rotOrder'
rdnet_in['ORIGINDESTINATION'] = 'Yes'
rdnet_in['PALLETQTY'] = '0'
rdnet_in['REFERENCECATEGORY'] = 'Sales'
rdnet_in['REFERENCEDOCUMENT'] = 'SalesOrder'
rdnet_in['ROADNETINTERNALSESSIONID'] = '35411'
rdnet_in['ROADNETREGIONID'] = le_code
rdnet_in['ROUTECODE'] = ''
rdnet_in['SECONDDRIVER'] = ''
rdnet_in['SECONDTRAILER'] = ''
rdnet_in['SEQUENCEDISTANCE'] = '.000000'
rdnet_in['SEQUENCENUMBER'] = '1'
rdnet_in['SEQUENCETRAVELTIME'] = '0'
rdnet_in['STATUS'] = 'Error'
rdnet_in['STOPTYPE'] = 'stpStop'
rdnet_in['TOTALDISTANCE'] = '.000000'
rdnet_in['TOTALROUTEDISTANCE'] = '.000000'
rdnet_in['TRUCKANDTRAILERASSIGNED'] = 'No'
rdnet_in['UNITID'] = ''
rdnet_in['STOPSERVICETIME'] = '720'
rdnet_in['TOTALSERVICETIME'] = '720'
rdnet_in['TOTALTRAVELTIME'] = '0'
rdnet_in['LINEREFID'] = rdnet_in['INVENTTRANSID']

# %%
# Defaults that are specific per LE

if le_code == "ZA1":
    rdnet_in['DESCRIPTION'] = 'BLOEM_PLAN'
    rdnet_in['FIRSTDRIVER'] = '825196'
    rdnet_in['FIRSTTRAILER'] = 'ST29PTAIL'
    rdnet_in['SHIPPINGCARRIER'] = '0'
    rdnet_in['VEHICLEID'] = 'TT4X2TAIL'

elif le_code == "NA1":
    rdnet_in['DESCRIPTION'] = 'Windhoek_PLAN'
    rdnet_in['FIRSTDRIVER'] = 'NA1-000002'
    #rdnet_in['FIRSTTRAILER'] = 'LD1001'
    rdnet_in['FIRSTTRAILER'] = 'TT1001'
    rdnet_in['SHIPPINGCARRIER'] = '0'
    #rdnet_in['VEHICLEID'] = 'LD1002'
    rdnet_in['VEHICLEID'] = 'TT1002'

elif le_code == "UG1":
    rdnet_in['DESCRIPTION'] = 'Rwenzori_PLAN'
    rdnet_in['FIRSTDRIVER'] = 'UG1-000001'
    rdnet_in['FIRSTTRAILER'] = 'TT1003'
    rdnet_in['SHIPPINGCARRIER'] = 'INTERNAL'
    rdnet_in['VEHICLEID'] = 'TT1004'    
    
elif le_code == "MZ1":
    rdnet_in['DESCRIPTION'] = 'Chimoio_PLAN'
    rdnet_in['FIRSTDRIVER'] = 'MZ1-000001'
    rdnet_in['FIRSTTRAILER'] = 'TT1002'
    rdnet_in['SHIPPINGCARRIER'] = '0'
    rdnet_in['VEHICLEID'] = 'TT1003'    
    
else:
    print("No valid legal entity")
    raise SystemExit("Terminating")

# %%
date_txt = input('Enter the dispatch date in the format yyyy-mm-dd: ')
date_dt = pd.to_datetime(date_txt)

# %%
rdnet_in['ROUTECOMPLETETIME'] = date_dt
rdnet_in['ROUTECOMPLETETIME'] = rdnet_in['ROUTECOMPLETETIME'].dt.normalize() + pd.Timedelta(days=0) + pd.Timedelta(hours=13) + pd.Timedelta(minutes=19)

rdnet_in['ROUTESTARTTIME'] = date_dt
rdnet_in['ROUTESTARTTIME'] = rdnet_in['ROUTESTARTTIME'].dt.normalize() + pd.Timedelta(days=0) + pd.Timedelta(hours=4) + pd.Timedelta(minutes=0)

rdnet_in['SCHEDULEDARRIVALDATETIME'] = date_dt
rdnet_in['SCHEDULEDARRIVALDATETIME'] = rdnet_in['SCHEDULEDARRIVALDATETIME'].dt.normalize() + pd.Timedelta(days=0) + pd.Timedelta(hours=12) + pd.Timedelta(minutes=59)

rdnet_in['SCHEDULEDDELIVERYDATETIME'] = date_dt
rdnet_in['SCHEDULEDDELIVERYDATETIME'] = rdnet_in['SCHEDULEDDELIVERYDATETIME'].dt.normalize() + pd.Timedelta(days=0)

rdnet_in['SCHEDULEDSHIPDATETIME'] = date_dt
rdnet_in['SCHEDULEDSHIPDATETIME'] = rdnet_in['SCHEDULEDSHIPDATETIME'].dt.normalize() + pd.Timedelta(days=0) + pd.Timedelta(hours=4) + pd.Timedelta(minutes=10)

rdnet_in['STOPARRIVALTIME'] = date_dt
rdnet_in['STOPARRIVALTIME'] = rdnet_in['STOPARRIVALTIME'].dt.normalize() + pd.Timedelta(days=0) + pd.Timedelta(hours=8) + pd.Timedelta(minutes=28)

# %% [markdown]
# ### Get customer master in order to get the postal code

# %%
print("Loading the customer master to obtain the zipcode")

if le_code == "ZA1":
    customers=pd.read_parquet("./data/customers/customers.parquet")

elif le_code == "NA1" or le_code == "UG1" or le_code == "MZ1":
    customers=pd.read_csv("./data/customers/NA1_UG1_MZ1_Export-Customersaddresses V3.csv")

else:
    print("No valid customer master file")
    raise SystemExit("Terminating")
    

customers_short = customers[['ADDRESSZIPCODE','CUSTOMERACCOUNT']].copy()
customers_short['ADDRESSZIPCODE'] = customers_short['ADDRESSZIPCODE'].fillna(0)
customers_short['ADDRESSZIPCODE'] = customers_short['ADDRESSZIPCODE'].astype(int)
customers_short['ADDRESSZIPCODE'] = customers_short['ADDRESSZIPCODE'].astype(str)

# %%
rdnet_in_bk=rdnet_in.copy()

# %%
rdnet_in = pd.merge(
    rdnet_in,
    customers_short,
    how="inner",
    on=None,
    left_on='STOPLOCATIONID',
    right_on='CUSTOMERACCOUNT',
    left_index=False,
    right_index=False,
    sort=True,
    suffixes=("_x", "_y"),
    copy=True,
    indicator=False,
    validate=None,
)

# %%
rdnet_in.rename(columns={'ADDRESSZIPCODE':'STOPPOSTALCODE'}, inplace=True)
rdnet_in.drop(columns={'CUSTOMERACCOUNT'}, inplace=True, axis=1)

# %% [markdown]
# ### Split file into n files, each with a different ORIGINLOCATIONID
# - Read csv file with ORIGINLOCATIONIDs
# - Generate a dataframe with the unique list of warehouses
# - Add a column (Group) to that dataframe where the warehouses are assigned into n groups
# - Generate n output files, each with ORIGINLOCATIONID as from the csv file

# %%
input('Remember to update data/roadnet/sessionIDs.csv with the list of session IDs.  This will determine how the Roadnet inbound file will be split.  Press "Enter" to continue. ')
sessionids = pd.read_csv('data/roadnet/sessionIDs.csv')

# %%
print("Splitting inbound files per sessionID")

# %%
sessionids_list = sessionids.values.tolist()

# %%
unique_warehouses_np = rdnet_in['ORIGINLOCATIONID'].unique()
unique_warehouses_df = pd.DataFrame(unique_warehouses_np)
number_of_warehouses = len(rdnet_in['ORIGINLOCATIONID'].unique())
number_of_sessionIDs = len(sessionids)

# %%
number_of_warehouses

# %%
if number_of_warehouses < number_of_sessionIDs:
    raise SystemExit("There are more sessionIDs than warehouses - some of the output files will have no lines")

# %%
rdnet_in['WH_route_combination'] = rdnet_in['ORIGINLOCATIONID'].astype(str) + rdnet_in['ROADNETROUTE'].astype(str)

# %%
x = rdnet_in.groupby(['WH_route_combination']).agg({'INVENTTRANSID': 'count'}).sort_values('INVENTTRANSID',ascending=False).reset_index().copy()

# %%
x['sessionID'] = ''
user = 0
user_ascending = True
for i in range(len(x)):
    if user_ascending == True and user < number_of_sessionIDs:
        user = user + 1

    if user_ascending == False and user <= number_of_sessionIDs:
        if user > 1:
            user = user - 1
        else:
            user = number_of_sessionIDs

    if user_ascending == True and user == number_of_sessionIDs:
        user_ascending = False
        user = number_of_sessionIDs           

    x.at[i, 'sessionID'] = str(sessionids_list[user-1][0])

# %%
x.drop(columns={'INVENTTRANSID'}, inplace=True, axis=1)

# %% [markdown]
# ### Merge back the warehouse-based split into the dataframe

# %%
len(rdnet_in)

# %%
rdnet_in = pd.merge(
    rdnet_in,
    x,
    how="inner",
    on=None,
    left_on='WH_route_combination',
    right_on='WH_route_combination',
    left_index=False,
    right_index=False,
    sort=True,
    suffixes=("_x", "_y"),
    copy=True,
    indicator=False,
    validate=None,
)

# %%
len(rdnet_in)

# %%
rdnet_in.drop(columns={'WH_route_combination'}, inplace=True, axis=1)
rdnet_in.rename(columns={'sessionID':'DYNAMICSRETRIEVALSESSIONID'}, inplace=True)

# %% [markdown]
# ### Generate CSV files

# %%
print("Writing CSV files per sessionID")
# Clean up the folder where the inbound files are to be stored
path = "data/roadnet/inbound/"
dir_list = os.listdir(path)

for i in range(0,len(dir_list)):
    xml_file = path+dir_list[i]
    os.remove(xml_file)  

# %%
for sessionID in sessionids_list:
    mask = (rdnet_in['DYNAMICSRETRIEVALSESSIONID'] == sessionID[0])
    df_temp = rdnet_in[mask].copy()
    filename = 'data/roadnet/inbound/rdnet_inbound_'+sessionID[0]
    df_temp.to_csv(filename+'.csv',index=False, lineterminator='\r\n')    
    df_temp.to_xml(filename+'.xml',index=False)    

# %%
print('The Roadnet order lines were split by warehouse and route ID, into ' + str(number_of_sessionIDs) + ' files.  The lines were split as follows:')
print(rdnet_in.groupby(['DYNAMICSRETRIEVALSESSIONID']).agg({'INVENTTRANSID': 'count'}))


# %% [markdown]
# ### End of script


