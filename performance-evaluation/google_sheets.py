# -*- coding: utf-8 -*-
from constants import google_credentials, GSHEETS_FILE_NAME, TEMPLATE_AUXILIAR_FILE_KEY, \
    TEMPLATE_TALENT_FILE_KEY, ANSWERS_ROLE_FILE_KEY
import pygsheets
import readchar


def create_sheet(title, folder_key):
    """With a given title and a key from a folder, create a Google Sheet in Drive"""
    print('Creando nuevo sheet...')
    sheet = google_credentials.create(title, template=None, folder=folder_key)
    print('Sheet ' + title + ' creada satisfactoriamente.\n')
    return sheet

def set_google_credentials():
    """Authenticate using OAuth from a given credentials file path and return credentials"""
    print('Autenticando...')
    global google_credentials
    google_credentials = pygsheets.authorize(GSHEETS_FILE_NAME)
    print('Autenticación exitosa!\n')

def get_sheet_by_key(key):
    """Given a key, open the element"""
    sheet = google_credentials.open_by_key(key)
    print('Documento ' + sheet.title + ' abierto.\n')
    return sheet

def get_feedback_sheet():
    """Ask for the document to read the feedback"""
    while True:
        feedback_file_key = input(
            'Ingresar la key del documento de feedback: ')
        try:
            feedback_sheet = google_credentials.open_by_key(feedback_file_key)
            print('Se va a utilizar el documento de feedback con el nombre: \'' +
                  feedback_sheet.title + '\'')
            print('Es correcto? Presione \'s/n\'.')
            if readchar.readchar() == 's':
                print('Abriendo documento para comenzar a trabajar...\n')
                return feedback_sheet
        except:
            print('El valor ingresado \'' + feedback_file_key +
                  '\' no es válido. Revisar y volver a intentar.\n')

def get_answers_role_sheet():
    """Open and return the answers role sheet"""
    print('Abriendo documento de respuestas de formulario para comenzar la copia...\n')
    answers_role_sheet = google_credentials.open_by_key(ANSWERS_ROLE_FILE_KEY)
    return answers_role_sheet

def get_answers_role_row(answers_role_sheet):
    """Ask for the row number in answers role sheet"""
    while True:
        answers_row_index = input('Ingresar número de fila correspondiente a la evaluación: ')
        try:
            print('Se van a copiar las respuestas para: \'' + answers_role_sheet.sheet1.cell('C' + answers_row_index).value + '\'')
            print('Es correcto? Presione \'s/n\'.')
            if readchar.readchar() == 's':
                print('Leyendo respuestas...')
                print('')
                return answers_row_index
        except:
            print('El valor ingresado \'' + answers_row_index + '\' no es válido. Revisar y volver a intentar.')
            print('')