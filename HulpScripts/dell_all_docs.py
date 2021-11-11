
import pandas as pd
from   MSGraphAPI import sharepoint_graph_API as AM_SP 
import logging
import datetime
import traceback
import os
import shutil
#import excelConfig as Exc_CNF

check = input("Alles deleten? type: 'ja': ")
if check == 'ja':
    # set a logging file 
    # Determine lof file name
    now = datetime.datetime.now()
    ts  = now.strftime('%Y-%m-%d-%H_%M_%S')
    #logfile
    #log_file_name= Exc_CNF.LOG_DIRECTORY+'\log_del_all_docs'+ts+'.txt'
    log_file_name= f'c:\TEMP\LOG\log_del_all_docs'+ts+'.txt'
    logging.basicConfig(  filename= log_file_name
                        #, encoding='utf-8'
                        , level=logging.DEBUG)


    #create SP object > based on config in SharepointConfig.py
    sp2 = AM_SP.SP_site()

    doc_list = sp2.list_doc_items_with_all_fields()
    for doc in doc_list:
        print('del docid: ' + str(doc['id']))
        sp2.del_list_item_on_DocID(doc['id'])
else:
     
    print('dan doen we niks')
