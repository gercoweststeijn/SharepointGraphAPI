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
SP_uploaded_Docs = sp.list_doc_items_with_all_fields()

print (SP_uploaded_Docs)


