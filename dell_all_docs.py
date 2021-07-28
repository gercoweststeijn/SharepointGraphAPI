
import pandas as pd
from   MSGraphAPI import SharepointGraphAPI as AM_SP 
import logging
import datetime
import traceback
import os
import shutil
import excelConfig as Exc_CNF

check = input("Alles deleten? type: 'ja'")
if input ==ja 'ja':
    # set a logging file 
    # Determine lof file name
    now = datetime.datetime.now()
    ts  = now.strftime('%Y-%m-%d-%H_%M_%S')
    #logfile
    log_file_name= Exc_CNF.BRON_DIRECTORY+'\log_del_all_docs'+ts+'.txt'
    logging.basicConfig(  filename= log_file_name
                        #, encoding='utf-8'
                        , level=logging.DEBUG)


    #create SP object > based on config in SharepointConfig.py
    sp2 = AM_SP.SP_site()

    doc_list = sp2.listDocItemsFields()
    for doc in doc_list:
        print('del docid: ' + str(doc['id']))
        sp2.delListItemonDocID(doc['id'])
else:
    print('dan doen we niks')
