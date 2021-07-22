#config file for sharepoint

TENANT_ID = '5d75e978-8a58-4197-a8d8-6eb786a5a8fa'
CLIENT_ID = 'ca88406e-cf1d-4945-9f82-ec7f59d08527'
SHAREPOINT_SITE = 'assets-docs-rwzi-db'

SHAREPOINT_HOST_NAME = 'wsaaenmaas.sharepoint.com'
AUTHORITY = 'https://login.microsoftonline.com/' + TENANT_ID
ENDPOINT = 'https://graph.microsoft.com/v1.0'
SCOPES = [
    'Files.ReadWrite.All',
    'Sites.ReadWrite.All',
    'User.Read',
    'User.ReadBasic.All'
]
DOC_LIST_TITLE = 'Documenten' # directory waar we de documenten inzetten

#to safeguad credentials
use_anonymous = True