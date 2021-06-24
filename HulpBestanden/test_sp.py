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

EXCEL_FILE = r'C:\Temp\Import test1.xlsx'
BLAD       = 'Blad1'
EXCEL_COL_NAME_TITEL = 'Titel'
EXCEL_COL_DOC_TYPE   = 'Documentsoort'
EXCEL_COL_OBJ_TYPE   = 'Objectsoort'
EXCEL_COL_UPLOADEN   = 'uploaden'

TENANT_ID = '5d75e978-8a58-4197-a8d8-6eb786a5a8fa'
CLIENT_ID = 'ca88406e-cf1d-4945-9f82-ec7f59d08527'
SHAREPOINT_SITE = 'assets-docs-rwzi-db'



sp = AM_SP.SP_site(   tenant_id      = TENANT_ID
                    , client_id      = CLIENT_ID
                    , sharepointsite = SHAREPOINT_SITE)

"""
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
"""

#docs = sp.listDocItemsFields()

#x = sp.updateStatusFinal()

#print (sp.access_token)

#result = sp.uploadFile(file = 'C:/Temp/log.txt', doc_id = 10, obj_id_list = [9])
print (sp.listDocItemsFields())
