"""SharePoint / Azure AD connection settings.

Fill in CLIENT_ID and TENANT_ID after registering an Azure AD app
(see README.md, "Set up SharePoint access" section). These two values are
not secrets by themselves (no client secret is used - this app signs users
in interactively), but they are specific to your organization's tenant.
"""

CLIENT_ID = "8fa12377-c129-4675-aade-8df44c75c6ab"
TENANT_ID = "9192e46f-bdc1-4c5e-85ff-a6b4b78d17fe"

# Derived from the SharePoint URL you shared:
# https://greenreb.sharepoint.com/sites/GreenrebExternalShareDrive/Shared Documents/General/Greenreb Shared External Marketing Drive
SHAREPOINT_HOSTNAME = "greenreb.sharepoint.com"
SHAREPOINT_SITE_PATH = "/sites/GreenrebExternalShareDrive"

# Everything under this folder (recursively, including all subfolders) gets
# indexed and searched - not the whole marketing drive, just this subtree.
SHAREPOINT_SEARCH_ROOT_PATH = "General/Greenreb Shared External Marketing Drive/Fotos & Videos"
