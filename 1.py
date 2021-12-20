from numpy.core.numeric import NaN
import pandas as pd
from   MSGraphAPI import sharepoint_graph_API as AM_SP 
import logging
import datetime
import traceback
import os
import shutil
import json

# haal lijst met lookupids op
def get_lookupids (list):
    ret_list = []
    for item in list:
        ret_list.append (item.get('LookupId'))
    return ret_list


sp = AM_SP.SP_site()


list = sp.get_SP_lists()

print (list)




"""
for doc in SP_uploaded_Docs:
    existing_file            =  doc['fields']['LinkFilename']
    print (existing_file)

    if 'td_documentsoortLookupId' in doc['fields']:
        existing_doc_type_id     =  doc['fields']['td_documentsoortLookupId']
    else:
        existing_doc_type_id = ''
    print (existing_doc_type_id)

    existing_deel_proces_value_list = get_lookupids(doc['fields']['td_deelproces_x002d_col'])   
    print (existing_deel_proces_value_list)

    if 'td_fabrikant' in doc['fields']:
        existing_fabrikant_value =  doc['fields']['td_fabrikant']
    else:
        existing_fabrikant_value = ''
    print (existing_fabrikant_value)

    #  if there is a location
    #        get list of location values 
    #  else the location is empty
    if 'td_locatie_x002d_col' in doc['fields']:
        existing_locatie_value_list = get_lookupids(doc['fields']['td_locatie_x002d_col'])
        new_row = pd.DataFrame(data = {"bestandsnaam" : existing_file, "documentsoort" : existing_doc_type_id,"deelproces" : existing_deel_proces_value_list,"fabrikant": existing_fabrikant_value,"locatie_list":  existing_locatie_value_list}) 
    
    else:
        new_row = pd.DataFrame(data = {"bestandsnaam" : existing_file, "documentsoort" : existing_doc_type_id,"deelproces" : existing_deel_proces_value_list,"fabrikant": existing_fabrikant_value}) #"locatie_list":  existing_locatie_value_list

    
    print('*************')
    new_row = pd.DataFrame(data = {"bestandsnaam" : existing_file, "documentsoort" : existing_doc_type_id,"deelproces" : existing_deel_proces_value_list,"fabrikant": existing_fabrikant_value}) #"locatie_list":  existing_locatie_value_list
    
    uploaded_docs_dataframe = uploaded_docs_dataframe.append(new_row, ignore_index=True)
"""