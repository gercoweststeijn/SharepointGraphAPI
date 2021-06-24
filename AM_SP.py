########################################################################################################################
# versie    auteur          toelichting
# 0.1       Weststeijn      Sharepoint class dat SP site via graph api ontsluit
#                           (https://docs.microsoft.com/en-us/graph/)
#
#                           Methoden toegeschreven naar het kunnen uploaden van bestanden en deze 'taggen'
#
#                          
# 0.2       Weststeijn     
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
    def __init__(self, tenant_id, client_id, sharepointsite):      

        self.TENANT_ID = tenant_id
        self.CLIENT_ID = client_id
        self.SHAREPOINT_SITE = sharepointsite

        ###############################################################################
        # deze instance variables zetten we hier hardcoded
        ###############################################################################
        self.SHAREPOINT_HOST_NAME = 'wsaaenmaas.sharepoint.com'
        self.AUTHORITY = 'https://login.microsoftonline.com/' + self.TENANT_ID
        self.ENDPOINT = 'https://graph.microsoft.com/v1.0'
        self.SCOPES = [
            'Files.ReadWrite.All',
            'Sites.ReadWrite.All',
            'User.Read',
            'User.ReadBasic.All'
        ]
        self.doc_list_title = 'Documenten' # directory waar we de documenten inzetten
        ##############################################################################


        #SP variables
        self.SP_URL = (f'https://{self.SHAREPOINT_HOST_NAME}/sites/{self.SHAREPOINT_SITE}')
        self.access_token = self.get_accesss_token()
        self.HEADERS={'Authorization': 'Bearer ' + self.access_token}

        #site variables
        self.site_id = self.get_SP_site_id()
        self.doc_lib_list_id = self.get_SP_doc_list_id()
        self.def_drive_id = self.get_def_drive_id()
        self.root_folder_id = self.def_drive_root_id()


   
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

    #return (configured) lists in SP
    def get_SP_lists(self):
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/lists/', headers=self.HEADERS)
        result.raise_for_status()
        sp_lists = result.json()
        return sp_lists['value']

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

    def get_doc_list(self):
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{self.doc_list_title}/items', headers=self.HEADERS)
        result.raise_for_status()
        doc_list = result.json()
        doc_list = doc_list['value']
        return doc_list


    # Upload a file 
    # 

    #
    def uploadFile(self, file):
        
        filename =  file [file.rfind('/')+1 : len(file)] 

        file_url = urllib.parse.quote(filename)
    
        # create upload session
        print (filename)
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

        file_status = os.stat(file)
        size = file_status.st_size
        CHUNK_SIZE = 10485760
        chunks = int(size / CHUNK_SIZE) + 1 if size % CHUNK_SIZE > 0 else 0

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
        upload_etag = upload_etag[0:pos_comma]
        return upload_etag
        
    
    def Etag2DocId(self, input_doc_etag):    
        #get all documents
        documents = self.get_doc_list()

        for document in documents:
            doc_etag = document['eTag'].upper()
            doc_etag = doc_etag[0:doc_etag.find(',')]
            if  doc_etag == input_doc_etag:
                doc_id = document['id']     
        return  doc_id              
           
    def updateDoctypeObjecttype (self, doc_id, doctype_id,objtype_id_list):
        update_instructions = {         
                                "td_documentsoortLookupId": doctype_id,
                                "td_objectsoortLookupId@odata.type": 'Collection(Edm.Int32)',
                                "td_objectsoortLookupId":  objtype_id_list,
                                } 

        result =  requests.patch(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{self.doc_lib_list_id}/items/{doc_id}/fields'
                        , headers=self.HEADERS
                        , json = update_instructions)

        result.raise_for_status()
        response = result.json()
        return 'Succes'

    def listDocItemsFields (self):
        result = requests.get(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{self.doc_list_title}/items/?expand=fields', headers=self.HEADERS)
        result.raise_for_status()
        doc_list = result.json()
        #doc_list = doc_list['value']
        return doc_list            

    def updateStatusFinal (self):
        update_instructions = {         
                                "_ModerationStatus": 0            
                              } 
                
        result =  requests.patch(f'{self.ENDPOINT}/sites/{self.site_id}/lists/{self.doc_lib_list_id}/items/148/fields'
                        , headers=self.HEADERS
                        , json = update_instructions)
        result.raise_for_status()
        return result
   