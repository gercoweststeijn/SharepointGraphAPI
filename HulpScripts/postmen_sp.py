import pandas as pd
from   MSGraphAPI import SharepointGraphAPI as AM_SP 
import logging
import datetime
import traceback
import os
import shutil


def ifnan(var, val):
  if var is None:
    return val
  return var

""""
# set a logging file 
# Determine lof file name
now = datetime.datetime.now()
ts  = now.strftime('%Y-%m-%d-%H_%M_%S')
#logfile
log_file_name= Exc_CNF.BRON_DIRECTORY+'\log_'+ts+'.txt'
logging.basicConfig(  filename= log_file_name
                    #, encoding='utf-8'
                    , level=logging.DEBUG)


# open a file to record results
result_file_name = Exc_CNF.BRON_DIRECTORY+'\ResultFile_'+ts+'.txt'
result_file = open(result_file_name, "a")
"""


#create SP object > based on config in SharepointConfig.py
sp2 = AM_SP.SP_site()
print()
print (sp2.get_accesss_token())
print()
#doc_list = sp2.list_doc_items_with_all_fields()
print('-------------------------')
item = sp2.get_doc_item_with_all_fields(doc_id = 4288)
print (item)
print('-------------------------')


lists = sp2.get_SP_lists()
# get de lijst ids voor documenten en 
for item in lists:
    print (item['name'])
    #if item['name'] ==  'TechnischDossierObjectsoorten':
    #    objSoortList_id = item['id']  
    #if item['name'] ==  'TechnischDossierDocumentsoorten':
    #    DocSoortList_id = item['id']
    if item['name'] ==  'TechnischDossierDeelprocessen':
        DeelprocesSoortList_id = item['id']
print('-------------------------')
dp_list = sp2.get_listDict_titleId(list_id = DeelprocesSoortList_id)

for item in dp_list:
  print (item)
print('-------------------------')
objDict = sp2.get_listDict_titleId(list_id = DeelprocesSoortList_id)
print (objDict)


