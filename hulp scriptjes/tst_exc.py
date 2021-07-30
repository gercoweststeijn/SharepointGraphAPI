#*********************************************************************************************************
#  versie    auteur          toelichting
#  0.1       Weststeijn      creatie
# 
#*********************************************************************************************************
# Dit script upload een set aan bestanden naar een sharepoint omgeving via microsoft graph apis
# (https://docs.microsoft.com/en-us/graph/)
# het doorloopt daarvoor de volgende stappen: 
# - 
# - 
# - 
# - 
# - 
#*********************************************************************************************************
# Benodigd
# - AM_SP.py = voor aanroepen van graph API
# Verder:
#     -pandas as pd
#     -AM_SP 
#     -logging
#     -datetime
#     -traceback
#     -os
###########################################################################################################
# config files 
#  * excelConfig.py
#  * SharepointConfig.py
#*********************************************************************************************************
# Authenticatie vind plaats via een external 
# device flow
# De flow maakt een url aan waarop een gebruiker 
# via de standaard auth. van AM kan authenticeren
#
# AUTH constants
#
#*********************************************************************************************************
#
#

import pandas as pd
from   MSGraphAPI import SharepointGraphAPI as AM_SP 
import logging
import datetime
import traceback
import os
import shutil
import excelConfig as Exc_CNF

# helper function to remove NaN inserts 
def ifnan(var, val):
  if (var is None) or (var.upper() == 'NaN'.upper() ):
    return val
  return var






#create SP object > based on config in SharepointConfig.py
sp = AM_SP.SP_site()

lists = sp.get_SP_lists()
# get de lijst ids voor documenten en 
for item in lists:
    if item['name'] ==  'TechnischDossierObjectsoorten':
        objSoortList_id = item['id']  
    if item['name'] ==  'TechnischDossierDocumentsoorten':
        DocSoortList_id = item['id']

docDict = sp.get_listDict_titleId(list_id = DocSoortList_id)
objDict = sp.get_listDict_titleId(list_id = objSoortList_id)

data = pd.read_excel (Exc_CNF.EXCEL_FILE, sheet_name = Exc_CNF.BLAD)
df = pd.DataFrame(data, columns= [Exc_CNF.EXCEL_COL_NAME_TITEL,
                                  Exc_CNF.EXCEL_COL_DOC_TYPE,
                                  Exc_CNF.EXCEL_COL_OBJ_TYPE,
                                  Exc_CNF.EXCEL_COL_UPLOADEN,
                                  Exc_CNF.EXCEL_COL_FABRIKANT,
                                  Exc_CNF.EXCEL_COL_LOCATIE ])


# uitlezen directory

file_dict = {}
directory = Exc_CNF.BRON_DIRECTORY
for file in os.listdir(directory):
    f_name = (file[0:file.index('.')])
    file_dict[f_name] = file


for index, row in df.iterrows():
    doctype_id=''
    objtype_id=''
    doc_file_name=''
    doc_file=''
    try:
        if row[Exc_CNF.EXCEL_COL_UPLOADEN] == 'JA':
            
            # determine values from excel
            doc_row_name = row[Exc_CNF.EXCEL_COL_DOC_TYPE]
            obj_row_name = row[Exc_CNF.EXCEL_COL_OBJ_TYPE]
            doc_file_name = row[Exc_CNF.EXCEL_COL_NAME_TITEL]
            doc_fabrikant_value = ifnan(row[Exc_CNF.EXCEL_COL_FABRIKANT],' ')
            doc_locatie_value = ifnan(row[Exc_CNF.EXCEL_COL_LOCATIE],' ')

            print (doc_fabrikant_value)

            
    except Exception as e:
        #print ('Jammer mislukt')
        #print (traceback.format_exc())
        result = 'Error: ' + repr(e)
        log_string = 'index: ' + str(index) + ' | ' + 'doctype_id: '+ str(doctype_id) + ' | ' +' objtype_id: '+ str(objtype_id) + ' | ' +' doc_file: '+ str(doc_file) + ' | ' + ' result: '+str(result) +'\n'



