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
#
#*********************************************************************************************************
# Authenticatie vind plaats via een external 
# device flow
# De flow maakt een url aan waarop een gebruiker 
# via de standaard auth. van AM kan authenticeren
#
# AUTH constants
#
#
#*********************************************************************************************************
#
#

import pandas as pd
import AM_SP 
import logging
import datetime
import traceback
import os

BRON_DIRECTORY = 'c:/Temp/test/WL'

EXCEL_FILE = r'C:\Temp\Import test3.xlsx'
BLAD       = 'Blad1'
EXCEL_COL_NAME_TITEL = 'Titel'
EXCEL_COL_DOC_TYPE   = 'Documentsoort'
EXCEL_COL_OBJ_TYPE   = 'Objectsoort'
EXCEL_COL_UPLOADEN   = 'uploaden'

TENANT_ID = '5d75e978-8a58-4197-a8d8-6eb786a5a8fa'
CLIENT_ID = 'ca88406e-cf1d-4945-9f82-ec7f59d08527'
SHAREPOINT_SITE = 'assets-docs-rwzi-db'


now = datetime.datetime.now()
ts  = now.strftime('%Y-%m-%d-%H_%M_%S')
log_file_name= BRON_DIRECTORY+'\log_'+ts+'.txt'
result_file_name = BRON_DIRECTORY+'\ResultFile_'+ts+'.txt'
result_file = open(result_file_name, "a")

logging.basicConfig(  filename= log_file_name
                    #, encoding='utf-8'
                    , level=logging.DEBUG)

sp = AM_SP.SP_site(   tenant_id      = TENANT_ID
                    , client_id      = CLIENT_ID
                    , sharepointsite = SHAREPOINT_SITE)

logging.info('Ophalen lijsten en ids')
lists = sp.get_SP_lists()
# get de lijst ids voor documenten en 
for item in lists:
    if item['name'] ==  'TechnischDossierObjectsoorten':
        objSoortList_id = item['id']  
    if item['name'] ==  'TechnischDossierDocumentsoorten':
        DocSoortList_id = item['id']

docDict = sp.get_listDict_titleId(list_id = DocSoortList_id)
objDict = sp.get_listDict_titleId(list_id = objSoortList_id)

logging.info('inlezen excel')
data = pd.read_excel (EXCEL_FILE, sheet_name = BLAD)
df = pd.DataFrame(data, columns= [EXCEL_COL_NAME_TITEL,EXCEL_COL_DOC_TYPE,EXCEL_COL_OBJ_TYPE,EXCEL_COL_UPLOADEN ])

# uitlezen directory 
logging.info('uitlezen')
file_dict = {}
directory = BRON_DIRECTORY
for file in os.listdir(directory):
    f_name = (file[0:file.index('.')])
    file_dict[f_name] = file

doctype_id=''
objtype_id=''
doc_file_name=''
doc_file=''


logging.info('loop excel')
for index, row in df.iterrows():
    try:
        if row[EXCEL_COL_UPLOADEN] == 'ja':
            
            doc_row_name = row[EXCEL_COL_DOC_TYPE]
            obj_row_name = row[EXCEL_COL_OBJ_TYPE]
            
            doctype_id = docDict[doc_row_name]
            objtype_id = objDict[obj_row_name]
            doc_file_name = row[EXCEL_COL_NAME_TITEL]
            doc_file = file_dict[doc_file_name]
 
            file_name = BRON_DIRECTORY+'/'+ str(doc_file)            
           
            # MS does not support directly linking lists to uploaded files
            # Therefore 
            #    * we determine the uploaded doc id based on th return etag
            #    * update the list values for the doc.
            doc_etag = sp.uploadFile(file = file_name)                       
            doc_id =  sp.Etag2DocId(input_doc_etag = doc_etag)            
            
            result = sp.updateDoctypeObjecttype(doc_id = doc_id, doctype_id = doctype_id, objtype_id_list = [objtype_id])

            log_string = 'index: ' + str(index) + ' | ' + 'doctype_id: '+ str(doctype_id) + ' | ' +' objtype_id: '+ str(objtype_id) + ' | ' +' doc_file: '+ str(doc_file) + ' | ' + ' result: '+str(result) +'\n'
            
            logging.info(log_string)
            result_file.write(log_string)
            print ('Weer een gelukt')

    except Exception as e:
        print ('Jammer mislukt')
        print (traceback.format_exc())
        result = 'Error: ' + repr(e)
        log_string = 'index: ' + str(index) + ' | ' + 'doctype_id: '+ str(doctype_id) + ' | ' +' objtype_id: '+ str(objtype_id) + ' | ' +' doc_file: '+ str(doc_file) + ' | ' + ' result: '+str(result) +'\n'
        logging.info(log_string)
        result_file.write(log_string)


result_file.close()
