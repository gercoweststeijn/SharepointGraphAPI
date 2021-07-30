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
  if (var is None) or (var == 'NaN' or 'nan' or 'NAN' ): # LELEIJK!!
    return val
  return var


lowerbounds = 3001
upperbounds = 4000

# set a logging file 
# Determine lof file name
now = datetime.datetime.now()
ts  = now.strftime('%Y-%m-%d-%H_%M_%S')
#logfile
log_file_name= Exc_CNF.LOG_DIRECTORY+'\log_'+str(lowerbounds)+'_'+str(upperbounds)+'___'+ts+'.txt'
logging.basicConfig(  filename= log_file_name
                    #, encoding='utf-8'
                    , level=logging.DEBUG)


# open a file to record results
result_file_name = Exc_CNF.LOG_DIRECTORY+'\ResultFile_'+str(lowerbounds)+'_'+str(upperbounds)+'___'+ts+'.txt'
result_file = open(result_file_name, "a", encoding='utf-8')

# open a file to record results
sweep_file_name = Exc_CNF.LOG_DIRECTORY+'\sweepFile_'+str(lowerbounds)+'_'+str(upperbounds)+'___'+ts+'.txt'
sweep_file = open(sweep_file_name, "a", encoding='utf-8')
sweep_file.write ('doc_etag'+ ','+'doctype_id' + ','+ 'objtype_id' + ','+ 'doc_fabrikant_value' + ','+ 'doc_locatie_value')

#create SP object > based on config in SharepointConfig.py
sp = AM_SP.SP_site()

logging.info('Ophalen sharepoint lijsten en ids')
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
data = pd.read_excel (Exc_CNF.EXCEL_FILE, sheet_name = Exc_CNF.BLAD)
df = pd.DataFrame(data, columns= [Exc_CNF.EXCEL_COL_NAME_TITEL,
                                  Exc_CNF.EXCEL_COL_DOC_TYPE,
                                  Exc_CNF.EXCEL_COL_OBJ_TYPE,
                                  Exc_CNF.EXCEL_COL_UPLOADEN,
                                  Exc_CNF.EXCEL_COL_FABRIKANT,
                                  Exc_CNF.EXCEL_COL_LOCATIE ])


# uitlezen directory
logging.info('uitlezen directory - samenstellen file dict')
file_dict = {}
directory = Exc_CNF.BRON_DIRECTORY
for file in os.listdir(directory):
    f_name = (file[0:file.index('.')])
    file_dict[f_name] = file

logging.info('loop excel')
for index, row in df.iterrows():
    doctype_id=''
    objtype_id=''
    doc_file_name=''
    doc_file=''
    doc_fabrikant_value = ''
    doc_locatie_value = ''
    try:
        
        if (row[Exc_CNF.EXCEL_COL_UPLOADEN] == 'JA') and (lowerbounds <= index <= upperbounds):
        #if row[Exc_CNF.EXCEL_COL_UPLOADEN]) == 'JA' :             
            # determine values from excel
            doc_row_name = row[Exc_CNF.EXCEL_COL_DOC_TYPE]
            obj_row_name = row[Exc_CNF.EXCEL_COL_OBJ_TYPE]
            doc_file_name = row[Exc_CNF.EXCEL_COL_NAME_TITEL]
            doc_fabrikant_value = ifnan(row[Exc_CNF.EXCEL_COL_FABRIKANT],' ')
            doc_locatie_value = ifnan(row[Exc_CNF.EXCEL_COL_LOCATIE],'')


            #<<  The following three statements may throw the (key value) error >>
            # determine doc and obj based on list values 
            doctype_id = docDict[doc_row_name]
            objtype_id = objDict[obj_row_name]            
            # determine file based on file dict
            doc_file = file_dict[doc_file_name]
            
            # format file - location
            file_name = Exc_CNF.BRON_DIRECTORY+'/'+ str(doc_file)            
            file_name_done = Exc_CNF.DONE_DIRECTORY+'/'+ str(doc_file)            
           
            # MS does not support directly linking lists to uploaded files
            # Therefore 
            #    * upload the file
            #    * we determine the uploaded doc id based on the returned etag
            #    * update the list values for the doc.            
            # 
            logging.info('uploaden bestand')
            doc_etag = sp.uploadFile(file = file_name)                       
            logging.info('Geupload met etag: '+str(doc_etag))

            #determine id of uploaded file 
            ret_doc_id =  sp.Etag2DocId(input_doc_etag = doc_etag)   
            if ret_doc_id == 0:
                result = 'ERROR: FOUTMELDING:  Doc id kon niet herleid worden van etag : '+str(doc_etag)
                sweep_data = (doc_etag+ ','+doctype_id + ','+ objtype_id + ','+ doc_fabrikant_value + ','+ doc_locatie_value +'\n')
                sweep_file.write (sweep_data)
            else: 
                # update the 
                result = sp.updateDoctypeObjecttypeFabrikantLocatie(doc_id = ret_doc_id, 
                                                                    doctype_id = doctype_id,  
                                                                    objtype_id_list = [objtype_id], 
                                                                    fabrikant_value = doc_fabrikant_value,
                                                                    locatie_value   = doc_locatie_value
                                                                )
                # move the file to move folder
                shutil.move(file_name, file_name_done)

            #create string to log
            log_string = 'index: ' + str(index) + ' | ' + \
                          ' doctype_id: '+ str(doctype_id) + ' | ' + \
                          ' objtype_id: '+ str(objtype_id) + ' | ' + \
                          ' doc_file: '+ str(doc_file) + ' | ' + \
                          ' fabrikant: ' + str(doc_fabrikant_value) + ' | '\
                          ' locatie: ' + str (doc_locatie_value)  + ' | '\
                          ' result: '+ str(result) + \
                          '\n'
            logging.info(log_string)
            result_file.write(log_string)
            #print ('Weer een gelukt')
            print (log_string)
            

    except Exception as e:
        #print ('Jammer mislukt')
        #print (traceback.format_exc())
        result = 'Error: ' + repr(e)
        log_string = 'index: ' + str(index) + ' | ' + \
                          ' doctype_id: '+ str(doctype_id) + ' | ' + \
                          ' objtype_id: '+ str(objtype_id) + ' | ' + \
                          ' doc_file: '+ str(doc_file) + ' | ' + \
                          ' fabrikant: ' + str(doc_fabrikant_value) + ' | '\
                          ' locatie: ' + str (doc_locatie_value)  + ' | '\
                          ' result: '+ str(result) + \
                          '\n'
        logging.info(log_string)
        logging.info ('TRACE:  '+traceback.format_exc())
        result_file.write(log_string)
        print (log_string)

# close log file
result_file.close()
sweep_file.close()