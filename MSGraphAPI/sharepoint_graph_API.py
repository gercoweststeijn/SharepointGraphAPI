########################################################################################################################
# versie    auteur          toelichting
# 0.1       Weststeijn      Sharepoint class dat SP site via graph api ontsluit
#                           (https://docs.microsoft.com/en-us/graph/)
#
#                           Methoden toegeschreven naar het kunnen uploaden van bestanden en deze 'taggen'
#
#                          
#
# TO DO                      -  exception handling op requests
#
#########################################################################################################################
#  
#
#  Authenticatie en autorisatie: 
# verloopt via een "external device flow". Functie maakt een url aan waarop een gebruiker via 'reguliere' wijze kan 
# authenticeren. opgehaalde token wordt gecashd. #
#
#  benodigdheden:  
#   Azure app waarvan bekend is: 
#      -TENANT_ID            
#      -CLIENT_ID            
#      -SHAREPOINT_HOST_NAME 
#########################################################################################################################


import requests
import msal
import atexit
import os.path
import urllib
import json



"""
Initialiseer een SP site en auth om een token en dus header te verkrijgen
"""
class SP_site:
    
    # constructor > 
    # authentiseert en bepaalt site waardes
    #
    def __init__(self):  
        # haal config data op
        self.read_config_file()
        
        #SP variables
        self.SP_URL = (f'https://{self.SHAREPOINT_HOST_NAME}/sites/{self.SHAREPOINT_SITE}')
        self.access_token = self.get_accesss_token()
        self.HEADERS={'Authorization': 'Bearer ' + self.access_token}

        #site variables
        self.site_id = self.get_SP_site_id()
        self.doc_lib_list_id = self.get_SP_doc_list_id()
        self.def_drive_id = self.get_def_drive_id()
        self.root_folder_id = self.def_drive_root_id()   

    def read_config_file(self):
        # read json file 
        config_file = open('MSGraphAPI/sharepoint_config.json',)
        config_data = json.load(config_file)
        
        # config data 
        self.TENANT_ID = config_data['TENANT_ID']
        self.CLIENT_ID = config_data['CLIENT_ID']
        self.SHAREPOINT_SITE = config_data['SHAREPOINT_SITE']
        self.SHAREPOINT_HOST_NAME = config_data['SHAREPOINT_HOST_NAME']
        self.AUTHORITY_BASE = config_data['AUTHORITY_BASE']
        self.AUTHORITY = self.AUTHORITY_BASE + self.TENANT_ID
        self.ENDPOINT = config_data['ENDPOINT']
        self.SCOPES = config_data['SCOPES']
        self.doc_list_title = config_data['DOC_LIST_TITLE']
                
        # Closing file
        config_file.close()

    """
    authentiseer en return een token
    Voor het authentiseren wordt een URL gegenereerd die de gebruiker moet volgen 
    om via standaard Aa en Maas authenticatie te authentiseren
    --
    tokens worden gecasht zodat voor een (op ?azure? ingerichte) periode niet geauthentiseerd hoeft te worden 
    """
    def get_accesss_token (self):
        cache = msal.SerializableTokenCache()

        if os.path.exists('token_cache.bin'):
            cache.deserialize(open('token_cache.bin', 'r').read())

        atexit.register(lambda: open('token_cache.bin', 'w').write(cache.serialize()) if cache.has_state_changed else None)

        app = msal.PublicClientApplication(self.CLIENT_ID, authority=self.AUTHORITY, token_cache=cache)

        accounts = app.get_accounts()
        result = None
        if len(accounts) > 0:
            result = app.acquire_token_silent(self.SCOPES, account=accounts[0])
            
        if result is None:
            flow = app.initiate_device_flow(scopes=self.SCOPES)
            if 'user_code' not in flow:
                raise Exception('Failed to create device flow')
            print(flow['message'])
            result = app.acquire_token_by_device_flow(flow)    
            token = result['access_token']
        else: 
            token = result['access_token']       
        return token

    # determine SP site ID
    def get_SP_site_id (self):
        result = requests.get(f'{self.ENDPOINT}/sites/{self.SHAREPOINT_HOST_NAME}:/sites/{self.SHAREPOINT_SITE}', headers=self.HEADERS)
        result.raise_for_status()
        site_info = result.json()
        site_id = site_info['id']
        return site_id

    # get SP documenten list id
    def get_SP_doc_list_id (self):
        list_title = 'Documenten'
        #get list by title
        #GET https://graph.microsoft.com/v1.0/sites/{site-id}/lists/{list-title}
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{list_title}', headers=self.HEADERS)
        result.raise_for_status()
        doc_library_list = result.json()
        doc_lib_list_id = doc_library_list['id']
        return doc_lib_list_id

    # get default drive_id (root)
    def get_def_drive_id(self):
        #get default drive 
        #GET https://graph.microsoft.com/v1.0//sites/{siteId}/drive
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/drive', headers=self.HEADERS)
        result.raise_for_status()
        def_drive = result.json()
        def_drive_id = def_drive['id']
        return def_drive_id

    # return root folder id (document lib.)
    def def_drive_root_id(self):
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/drives/{self.def_drive_id}/root', headers=self.HEADERS)
        result.raise_for_status()
        root_folder_info = result.json()
        root_folder_id = root_folder_info['id']
        return root_folder_id

    # retrun root folder children 
    def def_drive_root_children(self):
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/drives/{self.def_drive_id}/root/children', headers=self.HEADERS)
        result.raise_for_status()
        root_files_list = result.json()
        return root_files_list

    #return (configured) lists in SP
    def get_SP_lists(self):
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/lists/', headers=self.HEADERS)
        result.raise_for_status()
        sp_lists = result.json()
        sp_lists = sp_lists['value']
        return sp_lists

    def get_listDict_titleId(self, list_id):
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{list_id}/items?expand=fields', headers=self.HEADERS)
        result.raise_for_status()
        resp = result.json()

        list_dict = {'title': 'id'}
        for item in resp['value']:            
            title = item['fields']['LinkTitle']
            id    = item['id']
            list_dict[title] = id
        return list_dict    

    #
    # Upload a file to sharepoint
    # 
    def upload_file(self, file):
        
        filename =  file [file.rfind('/')+1 : len(file)] 

        file_url = urllib.parse.quote(filename)
    
        # create upload session
        result = requests.post(
            f'{self.ENDPOINT}/drives/{self.def_drive_id}/items/{self.root_folder_id}:/{file_url}:/createUploadSession',
            headers=self.HEADERS,
            json={
                '@microsoft.graph.conflictBehavior': 'replace',
                'description': '{filename}',
                'fileSystemInfo': {'@odata.type': 'microsoft.graph.fileSystemInfo'},
                'name': filename,
                "fields":{  
                        }
                } 
            )
        upload_session = result.json()        
        upload_url = upload_session['uploadUrl']

        # bepaal bestand en breek deze in stukjes om te uploaden
        file_status = os.stat(file)
        size = file_status.st_size
        CHUNK_SIZE = 10485760
        chunks = int(size / CHUNK_SIZE) + 1 if size % CHUNK_SIZE > 0 else 0

        # upload de losse stukjes bestand
        with open(file, 'rb') as fd:
            start = 0
            for chunk_num in range(chunks):
                chunk = fd.read(CHUNK_SIZE)
                bytes_read = len(chunk)
                upload_range = f'bytes {start}-{start + bytes_read - 1}/{size}'
                result = requests.put(
                    upload_url,
                    headers={
                        'Content-Length': str(bytes_read),
                        'Content-Range': upload_range
                    },
                    data=chunk
                )
                result.raise_for_status()
                start += bytes_read
            resp = result.json()
        
        # retrieve etag from created file and remove version (value after ,)
        # As version do not always match we remove these for later comparison
        upload_etag = resp['eTag'].upper()
        upload_etag = upload_etag.replace('{','')
        upload_etag = upload_etag.replace('}','')
        pos_comma =  (upload_etag.index(',') )
        upload_etag = upload_etag[1:pos_comma]
        return upload_etag
        
    
    def get_Etag_from_DocId(self, input_doc_etag):    
        #get all documents
        found_doc_id = 0
        documents = self.get_doc_list()
        for document in documents:
            doc_etag = document['eTag'].upper()
            doc_etag = doc_etag[1:doc_etag.find(',')]
            if  doc_etag == input_doc_etag.upper():
                found_doc_id = document['id']     
        return  found_doc_id        



    # Dedicated update function for RWZI DB sharepoint        
    #def update_doctype_objecttype_fabrikantLocatie (self, doc_id, doctype_id,objtype_id_list, fabrikant_value,locatie_value ):
    def update_doctype_objecttype_fabrikantLocatie (self, doc_id, doctype_id, fabrikant_value,locatie_value_list,deel_proces_value_list ):

        #
        # let op de hieronder gebruikte interne SP namen 
        # kunnen achterhaald worden in gui door de naar URL te kijken 
        # vanaf site settings > site collumns 
        # waarde staat in de url achter field=
        # https://wsaaenmaas.sharepoint.com/sites/assets-docs-rwzi-db/_layouts/15/mngfield.aspx
        # https://tomriha.com/what-is-sharepoint-column-internal-name-and-where-to-find-it/

        # reminder: https://stackoverflow.com/questions/65885995/microsoft-graph-api-patch-nested-structure
        update_instructions = { "td_fabrikant" : fabrikant_value,
                                "td_deelproces_x002d_colLookupId@odata.type": 'Collection(Edm.Int32)',
                                "td_deelproces_x002d_colLookupId": deel_proces_value_list,  
                                "td_documentsoortLookupId": doctype_id
                                 
                                } 
        if locatie_value_list != ['']:
            print ('Er is een locatie')
            update_instructions["td_locatie-colLookupId@odata.type"] =  'Collection(Edm.Int32)'
            update_instructions["td_locatie-colLookupId"]= locatie_value_list

        print ('instructie')
        print (update_instructions)                                                    
        print ('================')
        result =  requests.patch(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{self.doc_lib_list_id}/items/{doc_id}/fields'
        
                        , headers=self.HEADERS
                        , json = update_instructions)

        result.raise_for_status()
        response = result.json()
        return 'Succes'

    def list_doc_items_with_all_fields (self):
        nextpage = 0
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{self.doc_list_title}/items/?expand=fields', headers=self.HEADERS)
        result.raise_for_status()
        doc_list = result.json()
        if '@odata.nextLink' in doc_list:
            nextpage = 1 
            nexturl = doc_list['@odata.nextLink']
        doc_list = doc_list['value']

        while nextpage == 1:
            nextresult = requests.get(nexturl, headers=self.HEADERS)
            nextresult.raise_for_status()
            nextdoc_list = nextresult.json()
            if '@odata.nextLink' in nextdoc_list:
                nextpage = 1 
                nexturl = nextdoc_list['@odata.nextLink']
            else:
                nextpage = 0
            nextdoc_list = nextdoc_list['value']
            doc_list = doc_list + nextdoc_list
        
        return doc_list            

    def get_doc_item_with_all_fields (self, doc_id):
     
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{self.doc_list_title}/items/{doc_id}/?expand=fields', headers=self.HEADERS)
        result.raise_for_status()
        doc_fields = result.json()
        #doc_fields = doc_fields['value']
        return doc_fields

    def get_drive_itemid_from_doc_item (self, doc_id):
        # GET /sites/{site-id}/lists/{list-id}/items/{item-id}/driveItem
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{self.doc_list_title}/items/{doc_id}/driveItem', headers=self.HEADERS)
        result.raise_for_status()
        driveitem = result.json()
        drive_itemId = driveitem['id']
        return drive_itemId

    def get_doc_list(self):
        nextpage = 0
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{self.doc_list_title}/items', headers=self.HEADERS)
        result.raise_for_status()
        doc_list = result.json()
        if '@odata.nextLink' in doc_list:
            nextpage = 1 
            nexturl = doc_list['@odata.nextLink']

        doc_list = doc_list['value']

        while nextpage == 1:
            nextresult = requests.get(nexturl, headers=self.HEADERS)
            nextresult.raise_for_status()
            nextdoc_list = nextresult.json()
            if '@odata.nextLink' in nextdoc_list:
                nextpage = 1 
                nexturl = nextdoc_list['@odata.nextLink']
            else:
                nextpage = 0
            nextdoc_list = nextdoc_list['value']
            doc_list = doc_list + nextdoc_list   

        return doc_list

    def del_list_item_on_DocID (self, doc_id):
        result = requests.delete(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{self.doc_list_title}/items/{doc_id}', headers=self.HEADERS)
        result.raise_for_status()
        return result     

    #def delDriveItemonDocID (self, doc_id):
    #    #DELETE /sites/{siteId}/drive/items/{itemId}
    #    result = requests.delete(f'{self.ENDPOINT}/sites/{self.site_id}/drive/items/{doc_id}', headers=self.HEADERS)
    #    result.raise_for_status()
    #    return 'succes'
        
    # onderstaande werkt helaas niet vuia de graph api
    # # hiervoor gebruiken we eenzelfde call, maar dan via flow
    def update_status_to_Final (self):
        update_instructions = {         
                                "_ModerationStatus": 0            
                              } 
                
        result =  requests.patch(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{self.doc_lib_list_id}/items/148/fields'
                        , headers=self.HEADERS
                        , json = update_instructions)
        result.raise_for_status()
        return result
   