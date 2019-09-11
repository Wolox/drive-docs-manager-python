# -*- coding: utf-8 -*-
from constants import *
import pygsheets
import readchar


def create_sheet(google_credentials, title, folder_key):
    print('Creando nuevo sheet...')
    sheet = google_credentials.create(title, template=None, folder=folder_key)
    print('Sheet ' + title + ' creada satisfactoriamente.\n')
    return sheet

# Authenticate using OAuth from a given credentials file path and return credentials
def get_google_credentials():
    print('Autenticando...')
    google_credentials = pygsheets.authorize(GSHEETS_FILE_NAME)
    print('Autenticación exitosa!\n')
    return google_credentials

# Open and return the template sheets
def get_template_sheets(google_credentials):
    template_auxiliar_sheet = google_credentials.open_by_key(TEMPLATE_AUXILIAR_FILE_KEY)
    template_talent_sheet = google_credentials.open_by_key(TEMPLATE_TALENT_FILE_KEY)
    return (template_auxiliar_sheet, template_talent_sheet)

# Ask for the document to update and return the destiny sheet
def get_destiny_sheet(google_credentials, key):
    while True:
        try:
            destiny_sheet = google_credentials.open_by_key(key)
            print('Se va a actualizar el documento con el nombre: \'' + destiny_sheet.title + '\'')
            print('Es correcto? Presione \'s/n\'.')
            if readchar.readchar() == 's':
                print('Abriendo documento para comenzar a trabajar...\n')
                return destiny_sheet
        except:
            print('El valor ingresado \'' + destiny_file_key + '\' no es válido. Revisar y volver a intentar.\n')

# Ask for the document to read the auto evaluation
def get_auto_evaluation_sheet(google_credentials):
    while True:
        destiny_file_key = input('Ingresar la key del documento de auto evaluación: ')
        try:
            destiny_sheet = google_credentials.open_by_key(destiny_file_key)
            print('Se va a utilizar el documento de auto evaluación con el nombre: \'' + destiny_sheet.title + '\'')
            print('Es correcto? Presione \'s/n\'.')
            if readchar.readchar() == 's':
                print('Abriendo documento para comenzar a trabajar...\n')
                return destiny_sheet
        except:
            print('El valor ingresado \'' + destiny_file_key +
                  '\' no es válido. Revisar y volver a intentar.\n')

# Ask for the document to read the evaluation
def get_manager_evaluation_sheet(google_credentials):
    while True:
        destiny_file_key = input('Ingresar la key del documento de evaluación: ')
        try:
            destiny_sheet = google_credentials.open_by_key(destiny_file_key)
            print('Se va a utilizar el documento de evaluación con el nombre: \'' + destiny_sheet.title + '\'')
            print('Es correcto? Presione \'s/n\'.')
            if readchar.readchar() == 's':
                print('Abriendo documento para comenzar a trabajar...\n')
                return destiny_sheet
        except:
            print('El valor ingresado \'' + destiny_file_key + '\' no es válido. Revisar y volver a intentar.\n')

# Ask for the document to read the exchange evaluation
def get_exchange_evaluation_sheet(google_credentials):
    while True:
        destiny_file_key = input('Ingresar la key del documento de evaluación de intercambio: ')
        try:
            destiny_sheet = google_credentials.open_by_key(destiny_file_key)
            print('Se va a utilizar el documento de evaluación de intercambio con el nombre: \'' +
                  destiny_sheet.title + '\'')
            print('Es correcto? Presione \'s/n\'.')
            if readchar.readchar() == 's':
                print('Abriendo documento para comenzar a trabajar...\n')
                return destiny_sheet
        except:
            print('El valor ingresado \'' + destiny_file_key +
                  '\' no es válido. Revisar y volver a intentar.\n')

# Ask for the document to read the feedback
def get_feedback_sheet(google_credentials):
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

# Open and return the answers role sheet
def get_answers_role_sheet(google_credentials):
    print('Abriendo documento de respuestas de formulario para comenzar la copia...\n')
    answers_role_sheet = google_credentials.open_by_key(ANSWERS_ROLE_FILE_KEY)
    return answers_role_sheet
