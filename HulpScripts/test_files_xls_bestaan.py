#*********************************************************************************************************
#  versie    auteur          toelichting
#  0.1       Weststeijn      creatie
# 
# loop door excel om te checken of bestadnen in folder staan / aanwezig zijn
#
#*********************************************************************************************************

import pandas as pd
from   MSGraphAPI import sharepoint_graph_API as AM_SP 
import logging
import datetime
import traceback
import os
import shutil
import json


# helper function to remove NaN inserts 
# if var is None or nan return ''(=passed val) else return the var
def ifnan(var, val):
  if str(var) == ('NaN' or 'nan' or 'NAN' ) or (var is None) : # LELEIJK!!
    return val
  return var

def read_config_file():
    # stel globale variabelen vast 
    global BRON_DIRECTORY 
    global DONE_DIRECTORY
    global LOG_DIRECTORY
    global EXCEL_FILE 
    global BLAD 
    global EXCEL_COL_NAME_TITEL 
    global EXCEL_COL_DOC_TYPE
    #global EXCEL_COL_OBJ_TYPE 
    global EXCEL_COL_FABRIKANT 
    global EXCEL_COL_LOCATIE 
    global EXCEL_COL_UPLOADEN 
    global EXCEL_COL_DEELPROCES_TYPE
    global SP_LIJST_DOC_SOORT
    global SP_LIJST_COL_DEEL_PRO
    global SP_LIJST_COL_LOC

    # read json file 
    config_file = open('upload_config.json')
    config_data = json.load(config_file)
    
    # directories voor lezen / schrijven bestanden en log bestanden
    BRON_DIRECTORY = config_data['BRON_DIRECTORY']
    DONE_DIRECTORY = config_data['DONE_DIRECTORY']
    LOG_DIRECTORY = config_data['LOG_DIRECTORY']

    # Excel sheet configuratie
    # welke colom bevat welke data
    EXCEL_FILE = config_data['EXCEL_FILE']
    BLAD = config_data['BLAD']
    EXCEL_COL_NAME_TITEL = config_data['EXCEL_COL_NAME_TITEL']
    EXCEL_COL_DOC_TYPE = config_data['EXCEL_COL_DOC_TYPE']
    #EXCEL_COL_OBJ_TYPE = config_data['EXCEL_COL_OBJ_TYPE']
    EXCEL_COL_DEELPROCES_TYPE = config_data['EXCEL_COL_DEELPROCES_TYPE']
    EXCEL_COL_FABRIKANT = config_data['EXCEL_COL_FABRIKANT']
    EXCEL_COL_LOCATIE = config_data['EXCEL_COL_LOCATIE']
    EXCEL_COL_UPLOADEN = config_data['EXCEL_COL_UPLOADEN']        

    # ingerichte SP lijsten
    SP_LIJST_DOC_SOORT = config_data['SP_LIJST_DOC_SOORT']
    SP_LIJST_COL_DEEL_PRO = config_data['SP_LIJST_COL_DEEL_PRO']
    SP_LIJST_COL_LOC = config_data['SP_LIJST_COL_LOC']

    # Closing file
    config_file.close()


##############################################################################################
# START of script
##############################################################################################
# parameters om doorlezen van het excel te kunnen sturen
# lees regels tussen deze waarde


# haal config data op
read_config_file()

# set a logging file 
# Determine lof file name
now = datetime.datetime.now()
ts  = now.strftime('%Y-%m-%d-%H_%M_%S')


# open a file to record results
result_file_name = LOG_DIRECTORY+'bestanden_bestaan_'+ts+'.txt'
result_file = open(result_file_name, "a", encoding='utf-8')


data = pd.read_excel (EXCEL_FILE, sheet_name = BLAD)
excel_dataframe = pd.DataFrame(data, columns= [EXCEL_COL_NAME_TITEL,
                                               EXCEL_COL_DOC_TYPE,
                                               EXCEL_COL_UPLOADEN,
                                               EXCEL_COL_FABRIKANT,
                                               EXCEL_COL_LOCATIE, 
                                               EXCEL_COL_DEELPROCES_TYPE]
                              )

# uitlezen directory

file_dict = {}
directory = BRON_DIRECTORY

for file in os.listdir(directory):    
    # not all files seem to have and extension
    if '.' in file:
        f_name = (file[0:file.index('.')])
    else:
        f_name = file
    file_dict[f_name] = file

# we kunnen nu door het excel loopen 
logging.info('loop excel')
for index, row in excel_dataframe.iterrows():
    #clear values
    doctype_id=''
    objtype_id=''
    deelproces_id=''
    locatie_id=''
    exc_file_name=''
    exc_file=''
    exc_fabrikant_value = ''
    exc_locatie_value = ''
    exc_doc_row_name = ''
    exc_obj_row_name = ''
    
    if (row[EXCEL_COL_UPLOADEN] == 'JA') :
        

        # determine values from excel
#        exc_doc_row_name = row[EXCEL_COL_DOC_TYPE]
        exc_file_name = row[EXCEL_COL_NAME_TITEL]
        if exc_file_name in file_dict:
            result_file.write(str(exc_file_name) + '; JA' +'\n')
        else:
            result_file.write(str(exc_file_name) + '; NEE' +'\n')




    #         exc_deelproces_value = row[EXCEL_COL_DEELPROCES_TYPE]
    #         # check voor none / nan values
    #         # if there are any,  set them to an emtpy string
    #         if not pd.isnull(row[EXCEL_COL_FABRIKANT]): 
    #             exc_fabrikant_value = ifnan(row[EXCEL_COL_FABRIKANT],'')
    #         else: 
    #             exc_fabrikant_value = ''

    #         # check voor none / nan values
    #         if not pd.isnull(row[EXCEL_COL_LOCATIE]): 
    #             exc_locatie_value = ifnan(row[EXCEL_COL_LOCATIE],'')
    #         else: 
    #             exc_locatie_value = ''
            
    #         #<<  The following statements may throw the (key value) error >>
    #         # determine doc and obj based on list values 
    #         doctype_id    = docDict[exc_doc_row_name]  
    #         deelproces_id =  proDict[exc_deelproces_value]
    #         if exc_locatie_value != '':
    #             locatie_id    = locDict[exc_locatie_value]
    #         # determine file based on file dict
    #         
            
    #         # format file - location
    #         file_name = BRON_DIRECTORY+'/'+ str(exc_file)            
    #         file_name_done = DONE_DIRECTORY+'/'+ str(exc_file)   

    #         # MS does not support directly linking lists to uploaded files
    #         # Therefore 
    #         #    * upload the file
    #         #    * we determine the uploaded doc id based on the returned etag
    #         #    * update the list values for the doc.            
    #         # 
    #         logging.info('uploaden bestand')
    #         doc_etag = sp.upload_file(file = file_name)                       
    #         logging.info('Geupload met etag: '+str(doc_etag))

    #         #determine id of uploaded file 
    #         ret_doc_id =  sp.get_Etag_from_DocId(input_doc_etag = doc_etag)   
    #         if ret_doc_id == 0:
    #             result = 'ERROR: FOUTMELDING:  Doc id kon niet herleid worden van etag : '+str(doc_etag)
    #             sweep_data = (doc_etag+ ','+doctype_id + ','+ objtype_id + ','+ exc_fabrikant_value + ','+ exc_locatie_value +'\n')
    #             sweep_file.write (sweep_data)
    #         else: 
    #             #
    #             # we kunnen niet een lege locatielijst patchen! 
    #             # hiervoor een uitzondering maken
    #             drive_itemId = sp.get_drive_itemid_from_doc_item(doc_id = ret_doc_id)
    #             result = sp.update_doctype_objecttype_fabrikantLocatie( doc_id = ret_doc_id, 
    #                                                                     doctype_id = doctype_id,  
    #                                                                     #objtype_id_list = [objtype_id], 
    #                                                                     fabrikant_value = exc_fabrikant_value,
    #                                                                     locatie_value_list   = [locatie_id],
    #                                                                     deel_proces_value_list = [deelproces_id]
    #                                                                     )
    #             # move the file to done folder
    #             shutil.move(file_name, file_name_done)
            
    #         #create string to log
    #         log_string = 'index: ' + str(index) + ' | ' + \
    #                         ' doctype_id: '+ str(doctype_id) + ' | ' + \
    #                         ' objtype_id: '+ str(objtype_id) + ' | ' + \
    #                         ' DeelProces_type_id: '+ str(deelproces_id) + ' | ' + \
    #                         ' exc_file: '+ str(exc_file) + ' | ' + \
    #                         ' fabrikant: ' + str(exc_fabrikant_value) + ' | '\
    #                         ' locatie: ' + str (exc_locatie_value)  + ' | '\
    #                         ' result: '+ str(result) + \
    #                         '\n'
    #         logging.info(log_string)
    #         result_file.write(log_string)
    #         print (log_string)            
    
    # except Exception as e:
    #     result = 'Error: ' + repr(e)
    #     log_string = 'index: ' + str(index) + ' | ' + \
    #                         ' doctype_id: '+ str(doctype_id) + ' | ' + \
    #                         ' objtype_id: '+ str(objtype_id) + ' | ' + \
    #                         ' DeelProces_type_id: '+ str(deelproces_id) + ' | ' + \
    #                         ' exc_file: '+ str(exc_file) + ' | ' + \
    #                         ' fabrikant: ' + str(exc_fabrikant_value) + ' | '\
    #                         ' locatie: ' + str (exc_locatie_value)  + ' | '\
    #                         ' result: '+ str(result) + \
    #                         '\n'
    #     logging.info(log_string)
    #     logging.info ('TRACE:  '+traceback.format_exc())
    #     result_file.write(log_string)
    #     print (log_string)
    
# close log file
result_file.close()

