from constants import GDRIVE_FILE_NAME
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from percol import Percol
import readchar

# Select an option from an interactive list made with Percol
def select_from_list(list):
    candidates = [element[0] for element in list]
    with Percol(
            actions=[],
            descriptors={'stdin': "", 'stdout': "", 'stderr': ""},
            candidates=iter(candidates)) as p:
        p.loop()
    selected_element = p.model_candidate.get_selected_result()
    result = [element for element in list if element[0] == selected_element]
    return result

# Authenticate in Google
def google_authentication():
    gauth = GoogleAuth()
    if gauth.credentials is None:
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.LoadCredentialsFile(GDRIVE_FILE_NAME)
        gauth.Authorize()
        gauth.SaveCredentialsFile(GDRIVE_FILE_NAME)
    drive = drive_authentication(gauth)
    return drive

# Authenticate in Google Drive
def drive_authentication(gauth):
    drive = GoogleDrive(gauth)
    return drive

# Return the key from the element who want to search in Google Drive
def get_key_by_path(key):
    drive = google_authentication()
    folders = [('No lo encuentro', False)]
    print('Trayendo registros desde Google Drive...\n')
    file_list = drive.ListFile({'q': "'{}' in parents and trashed=false".format(key)}).GetList()
    for file1 in file_list:
        folders.append((file1['title'], file1['id']))
    while True:
        results = select_from_list(folders)
        print('\nElegiste: ' + results[0][0] + '.\nEs correcto? Presione \'s/n\'.')
        if readchar.readchar() == 's' and results[0][1]:
            return results[0][1]
        else:
            print('Ejecución finalizada.')
            exit()