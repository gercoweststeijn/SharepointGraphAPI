
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


x = sp2.Etag2DocId(input_doc_etag = "4480FD97-B216-49C8-9CE2-E2B436D68216" )
print (x)