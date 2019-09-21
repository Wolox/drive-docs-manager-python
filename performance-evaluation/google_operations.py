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
    PERFORMANCE_STUDY_FOLDER_KEY, FOLDER_TYPE, FOLDER_BOOLEAN, FOLDER_MIMETYPE

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

def create_file(file_type, title, parent_key):
    """With a given title and a key from a folder, create a Google Sheet in Drive"""
    print('Creando nueva ' + file_type + ' ...')
    if file_type == FOLDER_TYPE:
        file = pydrive_credentials.CreateFile({'title': title, 
            "parents":  [{"id": parent_key}], 
            "mimeType": FOLDER_MIMETYPE})
        file.Upload()
        new_file = file['id']
    else:
        new_file = pygsheets_credentials.create(title, template=None, folder=parent_key)
    print(file_type + ' ' + title + ' creada satisfactoriamente.\n')
    return new_file

def open_file(key):
    """Given a key, open the sheet"""
    sheet = pygsheets_credentials.open_by_key(key)
    print('Documento ' + sheet.title + ' abierto.\n')
    return sheet

def get_feedback_sheet():
    """Ask for the document to read the feedback"""
    while True:
        feedback_file_key = input(
            'Ingresar la key del documento de feedback: ')
        try:
            feedback_sheet = pygsheets_credentials.open_by_key(feedback_file_key)
            print('Se va a utilizar el documento de feedback con el nombre: \'' +
                  feedback_sheet.title + '\'')
            print('Es correcto? Presione \'s/n\'.')
            if readchar.readchar() == 's':
                print('Documento abierto.\n')
                return feedback_sheet
        except:
            print('El valor ingresado \'' + feedback_file_key +
                  '\' no es válido. Revisar y volver a intentar.\n')

def select_from_list(list):
    """Select an option from an interactive list made with Percol"""
    candidates = [element[0] for element in list]
    with Percol(
            actions=[],
            descriptors={'stdin': "", 'stdout': "", 'stderr': ""},
            candidates=iter(candidates)) as p:
        p.loop()
    selected_element = p.model_candidate.get_selected_result()
    result = [element for element in list if element[0] == selected_element]
    return result[0]

def list_files(key, create_opt=False):
    files = [('Crear una nueva carpeta..', FOLDER_BOOLEAN)] if create_opt else []
    elements = pydrive_credentials.ListFile({'q': "'{}' in parents and trashed=false".format(key)}).GetList()
    for element in elements:
        files.append((element['title'], element['id']))
    return files

def find_key(parent_key, create_opt=False):
    """Return the key from the element who want to search in Google Drive"""
    files = list_files(parent_key, create_opt)
    while True:
        result_file = select_from_list(files)
        if check_input(result_file[0]):
            return result_file[1]

def find_and_open_file(parent_key, name):
    input('Al hacer click en Enter, te listará los archivos de ' + name + ' para que selecciones:')
    key = find_key(parent_key, False)
    file = open_file(key)
    return file