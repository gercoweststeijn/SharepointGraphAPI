#*********************************************************************************************************
#  versie    auteur          toelichting
#  0.1       Weststeijn      creatie
#  0.2       Weststeijn      - SP inrichting aangepast van object naar deelproces - object vooralsnog 'uitgesterd'
#                            - waardes voor het bepalen van de upload (nu allemaal) naar config file verplaatst
#*********************************************************************************************************
# Dit script upload een set aan bestanden naar een sharepoint omgeving via microsoft graph apis
# (https://docs.microsoft.com/en-us/graph/)
# het doorloopt daarvoor de volgende stappen: 
# - haal config data
# - voer basis zaken uit: zet logging aan etc..
# - maak een SP object
# - haal SP lijsten op - waarmee we de docs gaan taggen
# - lees uit te lezen directory met te uploaden bestanden 
# - lees excel met te uploaden docs
# - loop door het excel 
#   - match: 
#       - bestandsnaam met bestanden in dir
#       - SP lijsten met excel tag
#   - upload doc
#   - tag bestand met lijst waarde
#
#

#*********************************************************************************************************
# Benodigd
# - Sharepoint_graph_api.py = voor aanroepen van graph API
# Verder:
#     -pandas as pd
#     -AM_SP 
#     -logging
#     -datetime
#     -traceback
#     -os
#     - shutil
#     - json
###########################################################################################################
# config files 
#  * upload_config.json: config voor dit script, oa met te gebruiken directories en excel bestand
#  * sharepoint_config.json: config voor gebruik SP 
#*********************************************************************************************************
# Authenticatie vind plaats via een external 
# device flow
# De flow maakt een url aan waarop een gebruiker 
# via de standaard auth. van AM kan authenticeren
#
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

def get_lookupids (list):
    ret_list = []
    for item in list:
        ret_list.append (item.get('LookupId'))
    return ret_list

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

# haal config data op
read_config_file()


# open a file to record results
result_file_name = LOG_DIRECTORY+'\ResultFile_NIETBESTAANDE_BESTANDEN'+'.txt'
result_file = open(result_file_name, "a", encoding='utf-8')

data = pd.read_excel (EXCEL_FILE, sheet_name = BLAD)
excel_dataframe = pd.DataFrame(data, columns= [EXCEL_COL_NAME_TITEL,
                                               EXCEL_COL_DOC_TYPE,
                                               EXCEL_COL_UPLOADEN,
                                               EXCEL_COL_FABRIKANT,
                                               EXCEL_COL_LOCATIE, 
                                               EXCEL_COL_DEELPROCES_TYPE]
                              )

file_dict = {}
directory = BRON_DIRECTORY
for file in os.listdir(directory):
    f_name = (file[0:file.index('.')])
    file_dict[f_name] = file

# we kunnen nu door het excel loopen 
logging.info('loop excel')
for index, row in excel_dataframe.iterrows():
    if (row[EXCEL_COL_UPLOADEN] == 'JA'):
        exc_file_name=''
        exc_file=''
        try:        
            
            exc_file_name = row[EXCEL_COL_NAME_TITEL]        
            exc_file = file_dict[exc_file_name]
            
        except Exception as e:
            result_file.write(str(exc_file_name) +'\n') 
            print (exc_file_name)
# close log file
result_file.close()

