
import pandas 
from   MSGraphAPI import SharepointGraphAPI as AM_SP 
import logging
import datetime
import traceback
import os
import shutil
import excelConfig as Exc_CNF


input_result_file = 'ResultFile_2001_2500___2021-07-28-18_52_44.txt'
path_input_result_file = Exc_CNF.LOG_DIRECTORY+input_result_file
retry_error = "  result: Error: KeyError('uploadUrl')"


# set a logging file 
# Determine lof file name
now = datetime.datetime.now()
ts  = now.strftime('%Y-%m-%d-%H_%M_%S')
#logfile
log_file_name= Exc_CNF.LOG_DIRECTORY+'\log_retry'+input_result_file+'___'+ts+'.txt'
logging.basicConfig(  filename= log_file_name
                    #, encoding='utf-8'
                    , level=logging.DEBUG)


# open a file to record results
result_file_name = Exc_CNF.LOG_DIRECTORY+'\ResultFile_retry'+input_result_file+'___'+ts+'.txt'
result_file = open(result_file_name, "a", encoding='utf-8')

df = pandas.read_csv(path_input_result_file, sep = '|', header = None)

#create SP object > based on config in SharepointConfig.py
sp = AM_SP.SP_site()        

#for file in os.listdir(directory):
#    f_name = (file[0:file.index('.')])
#    file_dict[f_name] = file

logging.info('loop excel')
for index, line in df.iterrows():
    
    doctype_id=''
    objtype_id=''
    doc_file_name=''
    doc_file=''
    doc_fabrikant_value = ''
    doc_locatie_value = ''
    try:
        
        if line[6] == retry_error:
        #if row[Exc_CNF.EXCEL_COL_UPLOADEN]) == 'JA' :             
            # determine values from excel
            oude_index = line[0]
            oude_index = (oude_index[oude_index.index(':')+1:len(oude_index)]).strip()

            doc_row_name = line[1]
            doc_row_name = (doc_row_name[doc_row_name.index(':')+1:len(doc_row_name)]).strip()
            
            obj_row_name = line[2]
            obj_row_name = (obj_row_name[obj_row_name.index(':')+1:len(obj_row_name)]).strip()

            doc_file_name = line[3]
            doc_file_name = (doc_file_name[doc_file_name.index(':')+1:len(doc_file_name)]).strip()

            doc_fabrikant_value = line[4]
            doc_fabrikant_value = (doc_fabrikant_value[doc_fabrikant_value.index(':')+1:len(doc_fabrikant_value)]).strip()
                
            doc_locatie_value = line[5]
            doc_locatie_value = (doc_locatie_value[doc_locatie_value.index(':')+1:len(doc_locatie_value)]).strip()
            
            #<<  The following three statements may throw the (key value) error >>
            # determine doc and obj based on list values 
            doctype_id = doc_row_name
            objtype_id = obj_row_name            
            # determine file based on file dict
            doc_file = doc_file_name
            
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
                         'oude index ' + oude_index + ' | ' + \
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
        log_string =  'index: ' + str(index) + ' | ' + \
                      'oude index ' + oude_index + ' | ' + \
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
