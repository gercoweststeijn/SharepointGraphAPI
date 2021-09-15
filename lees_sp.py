import pandas as pd
from   MSGraphAPI import sharepoint_graph_API as AM_SP 
import logging
import datetime
import traceback
import os
import shutil


def ifnan(var, val):
  if var is None:
    return val
  return var


SP_LIJST_DOC_SOORT = 'TechnischDossierDocumentsoorten'
SP_LIJST_COL_DEEL_PRO = 'TechnischDossierObjectsoorten'
SP_LIJST_COL_LOC = 'Locaties'


# Archief
# Gedeelde documenten
# Locaties
# PreservationHoldLibrary
# PageDiagnosticsResultList70B82488D89F4AAPDR
# SharePointHomeCacheList
# TechnischDossierDocumentsoorten
# TechnischDossierObjectsoorten
# TechnischDossierDeelprocessen
# Toegangsaanvragen

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

print(sp2.get_accesss_token())

#doc_list = sp2.list_doc_items_with_all_fields()

#item = sp2.get_doc_item_with_all_fields(doc_id = 4288)
#print (item)

#print (doc_lis)

#print (doc_list)

#counter = 0
#for doc in doc_list:
#  counter = counter +1
  #  print (doc['fields']['FileLeafRef'])
  #print(doc['id'])
  #print(doc['eTag'])

#print (counter)

# lists = sp2.get_SP_lists()
# # get de lijst ids voor documenten en 
# for item in lists:
#     print (item['name'])
#     #if item['name'] ==  'TechnischDossierObjectsoorten':
#     #    objSoortList_id = item['id']  
#     if item['name'] ==  'TechnischDossierDocumentsoorten':
#         DocSoortList_id = item['id']
#     #if item['name'] ==  'TechnischDossierDeelprocessen':
#     #    DeelprocesSoortList_id = item['id']

# # dp_list = sp2.get_listDict_titleId(list_id = DeelprocesSoortList_id)

# # for item in dp_list:
# #   print (item)


# lists = sp2.get_SP_lists()
# # get de lijst ids voor documenten en 
# for item in lists:
#     #if item['name'] ==  'TechnischDossierObjectsoorten':
#     #    objSoortList_id = item['id']  
#     if item['name'] ==  SP_LIJST_DOC_SOORT:
#         DocSoortList_id = item['id']
#     if item['name'] ==  SP_LIJST_COL_DEEL_PRO:
#         DeelprocesSoortList_id = item['id']
#     if item['name'] ==  SP_LIJST_COL_LOC:
#         LocatieList_id = item['id']    

# # docDict = sp2.get_listDict_titleId(list_id = DocSoortList_id)
# # for doc in docDict:
# #   print (doc)
# #   print ('======')
# # #objDict = sp2.get_listDict_titleId(list_id = objSoortList_id)
# # #print (objDict)

# print ('*******************************************************')


# procDict = sp2.get_listDict_titleId(list_id = DeelprocesSoortList_id)
# for proc in procDict:
#   print (proc)

# #objDict = sp2.get_listDict_titleId(list_id = objSoortList_id)
# #print (objDict)