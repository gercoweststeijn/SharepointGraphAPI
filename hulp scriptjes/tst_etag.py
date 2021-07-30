
import pandas as pd
from   MSGraphAPI import SharepointGraphAPI as AM_SP 
import logging
import datetime
import traceback
import os
import shutil
import excelConfig as Exc_CNF

def ifnan(var, val):
  if var is None:
    return val
  return var

""""
# set a logging file 
# Determine lof file name
now = datetime.datetime.now()
ts  = now.strftime('%Y-%m-%d-%H_%M_%S')
#logfile
log_file_name= Exc_CNF.BRON_DIRECTORY+'\log_'+ts+'.txt'
logging.basicConfig(  filename= log_file_name
                    #, encoding='utf-8'
                    , level=logging.DEBUG)


# open a file to record results
result_file_name = Exc_CNF.BRON_DIRECTORY+'\ResultFile_'+ts+'.txt'
result_file = open(result_file_name, "a")
"""


#create SP object > based on config in SharepointConfig.py
sp2 = AM_SP.SP_site()


x = sp2.Etag2DocId(input_doc_etag = "5CC3321B-862A-45A7-8503-30C51A0DFCAA" )
print (x)

doc_list = sp2.listDocItemsFields()
#print (doc_lis)


for doc in doc_list:
  name = (doc['fields']['FileLeafRef'])
  if name == 'Datasheet waterteller Doseerunits PE indikking_ Diehl Corona M.pdf':
    print   (doc['eTag'])  
    

