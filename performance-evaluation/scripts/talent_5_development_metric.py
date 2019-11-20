# -*- coding: utf-8 -*-
import readchar
import pygsheets
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from percol import Percol
from inputs import get_title
from validations import check_input
from constants import pygsheets_credentials, pydrive_credentials, GSHEETS_FILE_NAME, \
    GDRIVE_FILE_NAME, TEMPLATE_AUXILIAR_FILE_KEY, TEMPLATE_TALENT_FILE_KEY, ANSWERS_ROLE_FILE_KEY, \
    PERFORMANCE_REPORT_FOLDER_KEY, FOLDER_TYPE, FOLDER_BOOLEAN, FOLDER_MIMETYPE

def google_authentication():
    """Authenticate with pydrive & pygsheets"""
    print('Autenticando...')
    global pygsheets_credentials
    global pydrive_credentials
    pygsheets_credentials = pygsheets.authorize(GSHEETS_FILE_NAME)
    gauth = GoogleAuth()
    if gauth.credentials is None:
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.LoadCredentialsFile(GDRIVE_FILE_NAME)
        gauth.Authorize()
        gauth.SaveCredentialsFile(GDRIVE_FILE_NAME)
    pydrive_credentials = GoogleDrive(gauth)
    print('Autenticación exitosa!\n')

def open_file(key):
    """Given a key, open the sheet"""
    sheet = pygsheets_credentials.open_by_key(key)
    return sheet

def get_report(files, parent):
    filtered_files = []
    for file in files:
        if (file[0].endswith('/19') or file[0].endswith('/18')) and ("desempeño".upper() in file[0].upper()):
            return file

def list_files(key, name):
    result_files = []
    drive_files = pydrive_credentials.ListFile({'q': "'{}' in parents and trashed=false".format(key)}).GetList()
    for file in drive_files:
        result_files.append((file['title'], file['id']))
    return result_files

def get_last_development_worksheet(sheet):
#     import ipdb; ipdb.set_trace()
    worksheets = [ws for ws in sheet.worksheets() if not ws.hidden]
    last_development_worksheet = list(filter(lambda each: each.title.startswith('Desarrollo') and not 'RID' in each.title, worksheets))
    last_development_worksheet = last_development_worksheet[0] if last_development_worksheet else None
    return last_development_worksheet

google_authentication()
results = []
last_development_worksheet = []
dev5_cells = ('G39', 'H40')

performance_report_folders = list_files(PERFORMANCE_REPORT_FOLDER_KEY, 'root')
for folder in performance_report_folders:
    files = list_files(folder[1], folder[0])
    report = get_report(files, folder)
    if not report:
        continue
    sheet = open_file(report[1])
    last_development_worksheet = get_last_development_worksheet(sheet)
    if not last_development_worksheet:
        continue
    print((report[0],last_development_worksheet.get_values(start='G39', end='H40', returnas='matrix')))