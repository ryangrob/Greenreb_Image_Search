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

# Where app updates are published (make_update.py --upload) and where every
# app looks for them. Only the people who publish releases should be able to
# write here.
# TODO before rollout: move this to an internal site agreed with the CTO -
# the site above is shared with external partners.
UPDATE_SITE_PATH = SHAREPOINT_SITE_PATH
UPDATE_FOLDER_PATH = "General/Image Search Releases"

# Where colleagues can download the app - put into invitation emails.
# The GitHub repository is private, so this only works for people with
# access to it; a SharePoint link shared with Greenreb is the better home.
DOWNLOAD_URL = "https://github.com/ryangrob/Greenreb_Image_Search/releases/latest"

# Where the Support form in the app sends messages.
SUPPORT_EMAIL = "grobr1903@gmail.com"

# Where shared collections are kept: one small file each, listing the
# collection's name, its owner, the email addresses it is shared with and
# its photos. Everyone using the app needs to be able to read and write here.
# TODO before rollout: an internal site agreed with the CTO, so colleagues'
# email addresses are not stored on the externally shared drive.
SHARED_SITE_PATH = SHAREPOINT_SITE_PATH
SHARED_FOLDER_PATH = SHAREPOINT_SEARCH_ROOT_PATH
