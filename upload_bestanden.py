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

# 
# Read the config file and set the values as global variables 
def read_config_file():
    # stel globale variabelen vast ()
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
lowerbounds = 0
upperbounds = 999999999

# haal config data op
read_config_file()

# set log and result file
# set a logging file 
# Determine lof file name

# date / time for logging
now = datetime.datetime.now()
ts  = now.strftime('%Y-%m-%d-%H_%M_%S')
#logfile
log_file_name= LOG_DIRECTORY+'\log_'+str(lowerbounds)+'_'+str(upperbounds)+'___'+ts+'.txt'
logging.basicConfig(  filename= log_file_name
                    #, encoding='utf-8'
                    , level=logging.DEBUG)

# open a file to record results
result_file_name = LOG_DIRECTORY+'\ResultFile_'+str(lowerbounds)+'_'+str(upperbounds)+'___'+ts+'.txt'
result_file = open(result_file_name, "a", encoding='utf-8')

#create SP object > based on config in sharepoint_config.json
sp = AM_SP.SP_site()

# haal SP lijsten op
logging.info('Ophalen sharepoint lijsten en ids')
lists = sp.get_SP_lists()
# get de lijst ids voor documenten en lijsten. 

for item in lists:
    if item['name'] ==  SP_LIJST_DOC_SOORT:
        DocSoortList_id = item['id']
    if item['name'] ==  SP_LIJST_COL_DEEL_PRO:
        DeelprocesSoortList_id = item['id']
    if item['name'] ==  SP_LIJST_COL_LOC:
        LocatieList_id = item['id']    # stel SP_liST dictionaries vast 

# we uploaden de bestanden en voegen daar tags aan toe
# Deze tags staan in lijsten in SP 
# we maken hier eenmalig dicst van die lijsten (naam + ID)  van 
# zodat we niet een call hoeven doen voor elke upload / mutatie
docDict = sp.get_listDict_titleId(list_id = DocSoortList_id)
proDict = sp.get_listDict_titleId(list_id = DeelprocesSoortList_id)
locDict = sp.get_listDict_titleId(list_id = LocatieList_id)

# Lees het Excel in conform geconfigureerde waarden ('upload_config.json')
logging.info('inlezen excel')
data = pd.read_excel (EXCEL_FILE, sheet_name = BLAD)
excel_dataframe = pd.DataFrame(data, columns= [EXCEL_COL_NAME_TITEL,
                                               EXCEL_COL_DOC_TYPE,
                                               EXCEL_COL_UPLOADEN,
                                               EXCEL_COL_FABRIKANT,
                                               EXCEL_COL_LOCATIE, 
                                               EXCEL_COL_DEELPROCES_TYPE]
                              )

# uitlezen directory waar alle bestanden staan die we willen uploaden
# we maken hier een dict van om 
logging.info('uitlezen directory - samenstellen file dict')
file_dict = {}
directory = BRON_DIRECTORY
for file in os.listdir(directory):
    f_name = (file[0:file.index('.')])
    file_dict[f_name] = file

#******************************************************************************************************
# START LOOP DOOR EXCEL bestand - upload and patch acties
#******************************************************************************************************
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
    try:        
        if (row[EXCEL_COL_UPLOADEN] == 'JA') and (lowerbounds <= index <= upperbounds):
            
            # for some reason acquiring a refresh token does not work properly without re-initializing the object
            # after a hour acces is denied (401) and all operations fail
            # we reinitialize the SP object every 20 files
            if index % 20 == 0:
                sp = AM_SP.SP_site()

            # determine values from excel
            exc_doc_row_name = row[EXCEL_COL_DOC_TYPE]
            exc_file_name = row[EXCEL_COL_NAME_TITEL]
            exc_deelproces_value = row[EXCEL_COL_DEELPROCES_TYPE]

            # check voor none / nan values
            # if there are any,  set them to an emtpy string to prevent setting the value to 'nan' in SP
            # a value for fabrikant is not mandatory 
            if not pd.isnull(row[EXCEL_COL_FABRIKANT]): 
                exc_fabrikant_value = ifnan(row[EXCEL_COL_FABRIKANT],'')
            else: 
                exc_fabrikant_value = ''

            # check voor none / nan values
            # if there are any,  set them to an emtpy string to prevent setting the value to 'nan' in SP
            # a value for location is not mandatory 
            if not pd.isnull(row[EXCEL_COL_LOCATIE]): 
                exc_locatie_value = ifnan(row[EXCEL_COL_LOCATIE],'')
            else: 
                exc_locatie_value = ''
            
            #<<  The following statements may throw the (key value) error >>
            # we check to see if the excel tag is a tag that has been configured in SP
            doctype_id    = docDict[exc_doc_row_name]  
            deelproces_id =  proDict[exc_deelproces_value]
            if exc_locatie_value != '':
                locatie_id    = locDict[exc_locatie_value]
            # We check to see whether the excel file exists in the local-folder that should hold all 'to be uploaded files'
            exc_file = file_dict[exc_file_name]
            
            # bepaal of het document al op SP staat 
            existing_doc_id = sp.get_docid_on_filename(filename = exc_file)
            
            # 0 -> het document bestaat nog niet > we uploaden deze 
            if existing_doc_id == 0:

                # format file - location
                file_name = BRON_DIRECTORY+'/'+ str(exc_file)            
                file_name_done = DONE_DIRECTORY+'/'+ str(exc_file)   

                # MS does not support directly linking lists to uploaded files
                # Therefore 
                #    * upload the file
                #    * we determine the uploaded doc id based on the returned etag
                #    * update the list values for the doc.            
                # 
                logging.info('uploaden bestand')
                doc_etag = sp.upload_file(file = file_name)                       
                logging.info('Geupload met etag: '+str(doc_etag))

                #determine id of uploaded file 
                ret_doc_id =  sp.get_Etag_from_DocId(input_doc_etag = doc_etag)   
                if ret_doc_id == 0:
                    result = 'ERROR: FOUTMELDING:  Doc id kon niet herleid worden van etag : '+str(doc_etag)
                else: 
                    result = sp.update_doctype_objecttype_fabrikantLocatie( doc_id = ret_doc_id, 
                                                                            doctype_id = doctype_id,  
                                                                            #objtype_id_list = [objtype_id], 
                                                                            fabrikant_value = exc_fabrikant_value,
                                                                            locatie_value_list   = [locatie_id],
                                                                            deel_proces_value_list = [deelproces_id]
                                                                            )
                    
                
                #create string to log
                log_string = 'index: ' + str(index) + ' | ' + \
                                ' doctype_id: '+ str(doctype_id) + ' | ' + \
                                ' DeelProces_type_id: '+ str(deel_proces_value_list) + ' | ' + \
                                ' exc_file: '+ str(exc_file_name) + ' | ' + \
                                ' fabrikant: ' + str(exc_fabrikant_value) + ' | '\
                                ' locatie: ' + str (locatie_value_list)  + ' | '\
                                ' result: '+ str(result) + \
                                '\n'
                logging.info(log_string)
                result_file.write(log_string)
                # we printen de voortgang 
                print (log_string)    
            
            # het document bestaat al wel op SP  -> we halen deze op en patchen 
            else:
                
                logging.info('Bestand staat al op sharepoint we gaan patchen')

                # haal de tags voor het sp document op
                doc_item_with_fields = sp.get_doc_item_with_all_fields(doc_id= existing_doc_id)
                #
                # We bepalen per tag of de excel waarde nieuw is of dat deze al bestaat
                # Fabrikant en doc type kan maar 1 waarde hebben -> als het er meer zijn geeft dit een fout 
                #   locatie en deelproces kunnen meerdere waarden hebben -> als dit zo is dan moeten we deze toevoegen en het bestand patchen
                #       Als dit echter niet zo is dan patchen we het bestand niet
                #
                #
                # NB fabrikant en locatie hoeven (nog) niet gevuld te zijn, dit checken we dus eerst
                # 

                # haal waarden op
                existing_doc_type = doc_item_with_fields['fields']['td_documentsoortLookupId']

                # bepaal of er een fabr waarde is
                if 'td_fabrikant' in doc_item_with_fields['fields']:
                    existing_fabrikant_value =  doc_item_with_fields['fields']['td_fabrikant']
                else:
                    existing_fabrikant_value = ''
                
                # match de doc types uit excel en SP
                if  existing_doc_type != doctype_id:
                    doc_type_match = False
                    result = 'Doc type matcht niet met bestaande doctype:'+str(existing_doc_id)
                else:
                    doc_type_match = True
                #match de fabr waarde uit excel en SP
                if  existing_fabrikant_value != exc_fabrikant_value:
                    fabrikant_value_match = True
                    result = result + 'fabrikant value matcht niet met bestaande fabrikant value:' + existing_fabrikant_value
                else:
                    fabrikant_value_match = True

                # als de fabr. en doc type matchen dan gaan we verder, zo niet dan schrijven we een foutmelding, via de gezette result waardes
                if fabrikant_value_match and doc_type_match:
                    # we checken of we moeten gaan updaten, we beginnen met de uitgaans waarde (false) dat dit niet zo is
                    updaten = False
                    locatie_value_list = get_lookupids(doc_item_with_fields['fields']['td_locatie_x002d_col'])


                    # Checken of  er een nieuwe locatie of deelproces waarde gegeven is >> dan moeten we het bestand gaan updaten
                    if locatie_id != '' and int(locatie_id) not in locatie_value_list:
                       locatie_value_list.append(int(locatie_id))
                       updaten = True
                    deel_proces_value_list = get_lookupids(doc_item_with_fields['fields']['td_deelproces_x002d_col'])
                    if deelproces_id != '' and int(deelproces_id) not in deel_proces_value_list:
                        deel_proces_value_list.append(int(deelproces_id))
                        updaten = True

                    #
                    # we updaten het bestand of geven aan dat dit niet nodig is.
                    #
                    if updaten:
                        # update drive item met geupdate waarden
                        result = sp.update_doctype_objecttype_fabrikantLocatie( doc_id = existing_doc_id, 
                                                                                doctype_id = doctype_id,  
                                                                                #objtype_id_list = [objtype_id], 
                                                                                fabrikant_value = exc_fabrikant_value,
                                                                                locatie_value_list   = locatie_value_list,
                                                                                deel_proces_value_list = deel_proces_value_list
                                                                                )

                        result = result + ' Bijgewerkt'
                        log_string = 'index: ' + str(index) + ' | ' + \
                                ' doctype_id: '+ str(doctype_id) + ' | ' + \
                                ' DeelProces_type_id: '+ str(deel_proces_value_list) + ' | ' + \
                                ' exc_file: '+ str(exc_file_name) + ' | ' + \
                                ' fabrikant: ' + str(exc_fabrikant_value) + ' | '\
                                ' locatie: ' + str (locatie_value_list)  + ' | '\
                                ' result: '+ str(result) + \
                                '\n'
                        logging.info(log_string)
                        result_file.write(log_string)
                        print (log_string)    
                    else: 
                        result = 'Updaten niet nodig deze locatie en deelproces waarde zijn al ingevuld '    
                        log_string = 'index: ' + str(index) + ' | ' + \
                                    ' doctype_id: '+ str(doctype_id) + ' | ' + \
                                    ' DeelProces_type_id: '+ str(deel_proces_value_list) + ' | ' + \
                                    ' exc_file: '+ str(exc_file_name) + ' | ' + \
                                    ' fabrikant: ' + str(exc_fabrikant_value) + ' | '\
                                    ' locatie: ' + str (locatie_value_list)  + ' | '\
                                    ' result: '+ str(result) + \
                                    '\n'
                        logging.info(log_string)
                        result_file.write(log_string)
                        print (log_string)   
                else:
                    # print/log de fout >> dat er geen match is 
                    log_string = 'index: ' + str(index) + ' | ' + \
                                ' doctype_id: '+ str(doctype_id) + ' | ' + \
                                ' DeelProces_type_id: '+ str(deel_proces_value_list) + ' | ' + \
                                ' exc_file: '+ str(exc_file_name) + ' | ' + \
                                ' fabrikant: ' + str(exc_fabrikant_value) + ' | '\
                                ' locatie: ' + str (locatie_value_list)  + ' | '\
                                ' result: '+ str(result) + \
                                '\n'
                    logging.info(log_string)
                    result_file.write(log_string)
                    print (log_string)  


    except Exception as e:
        result = 'Error: ' + repr(e)
        log_string = 'index: ' + str(index) + ' | ' + \
                    ' doctype_id: '+ str(doctype_id) + ' | ' + \
                    ' DeelProces_type_id: '+ str(deel_proces_value_list) + ' | ' + \
                    ' exc_file: '+ str(exc_file_name) + ' | ' + \
                    ' fabrikant: ' + str(exc_fabrikant_value) + ' | '\
                    ' locatie: ' + str (locatie_value_list)  + ' | '\
                    ' result: '+ str(result) + \
                    '\n'
        logging.info(log_string)
        logging.info ('TRACE:  '+traceback.format_exc())
        result_file.write(log_string)
        print (log_string)
    
# close log file
result_file.close()

