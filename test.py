import pandas as pd
import SharepointGraph as AM_SP 
import logging
import datetime
import traceback
import os
import excelConfig as Exc_CNF


sp = AM_SP.SP_site()

logging.info('Ophalen sharepoint lijsten en ids')
lists = sp.get_SP_lists()

for item in lists:
    print (item['name'])
    """"
    if item['name'] ==  'TechnischDossierObjectsoorten':
        objSoortList_id = item['id']  
    if item['name'] ==  'TechnischDossierDocumentsoorten':
        DocSoortList_id = item['id']
        """
docs  = sp.listDocItemsFields()

print (docs)

for doc in docs:
    if (doc['id']) == '382':
        print (doc)