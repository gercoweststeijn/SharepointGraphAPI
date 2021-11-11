#*********************************************************************************************************
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


def get_lookupids (list):
    ret_list = []
    for item in list:
        ret_list.append (item.get('LookupId'))
    return ret_list


#create SP object > based on config in sharepoint_config.json
sp = AM_SP.SP_site()

#print (sp.get_doc_items())

#print (sp.get_doc_item_with_all_fields(doc_id = 9577))
print (sp.access_token)
#print (sp.get_docid_on_name())

#print (sp.get_def_drive_id())
#A = sp.def_drive_root_children()

#print (get_docid_on_filename(filename = ''))
#print (sp.get_docid_on_filename(filename = 'Krachtinstallatie bergingsruimte 1.docx'))

#l= sp.get_doc_item_with_all_fields(doc_id=9379)

#ll  = (l['fields']['td_deelproces_x002d_col'])
#print (ll)

#ll = [{'LookupId': 123, 'LookupValue': '1207_CH4,H2S, en NH3 detectie'}, {'LookupId': 1234, 'LookupValue': '1207_CH4,H2S, en NH3 detectie'}]

#luv = get_lookupids(ll)
#print (luv)