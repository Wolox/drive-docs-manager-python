# -*- coding: utf-8 -*-
pygsheets_credentials = 0
pydrive_credentials = 0

# Mode for running script for ever evaluation after the first one.
NEXT_EVALUATION = "NEXT_EVALUATION"
# Mode for creating a RID evaluation.
RID_EVALUATION = "RID_EVALUATION"
# Mode for creating the auto evaluation form.
AUTO_EVALUATION = "AUTO_EVALUATION"
# Mode for creating the manager evaluation form.
MANAGER_EVALUATION = "MANAGER_EVALUATION"
# Mode for creating the exchage evaluation form.
EXCHANGE_EVALUATION = "EXCHANGE_EVALUATION"
# Mode for creating the first evaluation form after filling the agreement form.
FIRST_EVALUATION = "FIRST_EVALUATION"
# Mode for updating feedback for a given evaluation.
UPDATE_FEEDBACK = "UPDATE_FEEDBACK"

# Path to credentials file from Pygsheets. Read documentation https://pygsheets.readthedocs.io/en/latest/authorizing.html to create new credentials
GSHEETS_FILE_NAME = 'client_secret.json'

# Path to credentials file from PyDrive. Read documentation https://pythonhosted.org/PyDrive/quickstart.html to create new credentials
GDRIVE_FILE_NAME = 'credentials.json'

# For file named: 'Template - Evaluación Auxiliares'
TEMPLATE_AUXILIAR_FILE_KEY = '1eB-6j0xc9qeFVtuF3xXfajiWD9e_TcFdxxTHHdlAjyU'

# For file named: 'Template - Evaluación Talentos'
TEMPLATE_TALENT_FILE_KEY = '1D04Q-IQ67F1wTAgk3oEr1Q902DJd4ulv666mBlcar6c'

# For file named: 'Rol Laboral - Evaluaciones de Desempeño (Respuestas)'
ANSWERS_ROLE_FILE_KEY = '1felT_0RAVlG4FWFTbCkx3XMJjVSTd5sXqOLVYMzcRSo'

# Types of files i can create in Drive
FOLDER_BOOLEAN = "create#folder"
FOLDER_TYPE = "carpeta"
SHEET_TYPE = "sheet"
FOLDER_MIMETYPE = "application/vnd.google-apps.folder"
AUTO_EVALUATION_NAME = "autoevaluación"
MANAGER_EVALUATION_NAME = "evaluación"
EXCHANGE_EVALUATION_NAME = "intercambio"

OPERATION_CREATE_SHEET = "OPERATION_CREATE_SHEET"
OPERATION_COPY_TABS = "OPERATION_COPY_TABS"
OPERATION_BUILD_EVALUATION_FORM = "OPERATION_BUILD_EVALUATION_FORM"
OPERATION_HIDE_TALENTS = "OPERATION_HIDE_TALENTS"
OPERATION_COPY_ANSWERS = "OPERATION_COPY_ANSWERS"
OPERATION_COPY_FEEDBACK = "OPERATION_COPY_FEEDBACK"

AUTO_EVALUATION_FOLDER_KEY = "126gcCwdBklMyEnIgdYWSPEiwSD_jy-Z7"
MANAGER_EVALUATION_FOLDER_KEY = "1posz9Fkla1VZexbwQP32q0xo5aaDBSox"
EXCHANGE_EVALUATION_FOLDER_KEY = "1nLjfe7kcnsRHRof2kFz06DAmcMzK2Kfr"
RID_EVALUATION_FOLDER_KEY = "1hHBXRBU2ls5CVcHHfN4lzlNh6S20JeIS"
PERFORMANCE_STUDY_FOLDER_KEY = "1wGrZ9n-il5YnaS6Z_ihapCg1z81ad5Bn"
PERFORMANCE_REPORT_FOLDER_KEY = "0B-gOApkHyMPvelYzM3VGV2ZlYVk"

operations_by_mode_dictionary = {
    OPERATION_CREATE_SHEET:             [AUTO_EVALUATION, MANAGER_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION],
    OPERATION_COPY_TABS: 				[NEXT_EVALUATION, AUTO_EVALUATION, MANAGER_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION, RID_EVALUATION],
    OPERATION_BUILD_EVALUATION_FORM:	[EXCHANGE_EVALUATION, FIRST_EVALUATION],
    OPERATION_HIDE_TALENTS: 			[NEXT_EVALUATION, AUTO_EVALUATION, MANAGER_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION],
    OPERATION_COPY_ANSWERS: 			[NEXT_EVALUATION, AUTO_EVALUATION, MANAGER_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION],
    OPERATION_COPY_FEEDBACK:			[NEXT_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION, UPDATE_FEEDBACK]
}

# This is a matching between an action in google and a 2-tuple with:
# (Where I want to create the file, Where is the parent where I want to create the file)
file_actions = {
    NEXT_EVALUATION: {
        'create': False,
        'search': {
            'key': PERFORMANCE_REPORT_FOLDER_KEY
        }
    },
    AUTO_EVALUATION: {
        'create': {
            'key': AUTO_EVALUATION_FOLDER_KEY,
            'parent_key': PERFORMANCE_STUDY_FOLDER_KEY
        },
        'search': False
    },
    MANAGER_EVALUATION: {
        'create': {
            'key': MANAGER_EVALUATION_FOLDER_KEY,
            'parent_key': PERFORMANCE_STUDY_FOLDER_KEY
        },
        'search': False
    },
    EXCHANGE_EVALUATION: {
        'create': {
            'key': EXCHANGE_EVALUATION_FOLDER_KEY,
            'parent_key': PERFORMANCE_STUDY_FOLDER_KEY
        },
        'search': False
    },
    FIRST_EVALUATION: {
        'create': {
            'key': False,
            'parent_key': PERFORMANCE_REPORT_FOLDER_KEY
        },
        'search': False
    },
    RID_EVALUATION: {
        'create': {
            'key': RID_EVALUATION_FOLDER_KEY,
            'parent_key': PERFORMANCE_STUDY_FOLDER_KEY
        },
        'search': False
    }
    
}

folders_dictionary = {
    NEXT_EVALUATION: {
        'need_child_key': True,
        'parent_key': False,
        'parent_path': PERFORMANCE_REPORT_FOLDER_KEY
    },
    AUTO_EVALUATION: {
        'need_child_key': False,
        'parent_key': AUTO_EVALUATION_FOLDER_KEY,
        'parent_path': PERFORMANCE_STUDY_FOLDER_KEY
    },
    MANAGER_EVALUATION: {
        'need_child_key': False,
        'parent_key': MANAGER_EVALUATION_FOLDER_KEY,
        'parent_path': PERFORMANCE_STUDY_FOLDER_KEY
    },
    EXCHANGE_EVALUATION: {
        'need_child_key': False,
        'parent_key': EXCHANGE_EVALUATION_FOLDER_KEY,
        'parent_path': PERFORMANCE_STUDY_FOLDER_KEY
    },
    FIRST_EVALUATION: {
        'need_child_key': False,
        'parent_key': False,
        'parent_path': PERFORMANCE_REPORT_FOLDER_KEY
    },
    RID_EVALUATION: {
        'need_child_key': False,
        'parent_key': RID_EVALUATION_FOLDER_KEY,
        'parent_path': PERFORMANCE_STUDY_FOLDER_KEY
    }
}

# Structs declarations

# This is a matching between a talent and a 4-tuple with:
# (talent_identifier, columnt for talent in answers_role_sheet, row for talent in 'Desempeño', row in tab for title talents)
template_talents_dictionary = {
    'Universales': 					('U',	'J',	28,		[2, 12, 23, 35, 45, 56, 67, 78, 88, 99]),
    'Administración y Finanzas': 	('AF',	'O',	56,		[2, 12, 22]),
    'Business Dev': 				('BD',	'Q',	72,		[2, 12, 22, 32, 42, 52]),
    'Calidad': 						('C',	'T',	76,		[2, 12, 22, 32, 42]),
    'Desarrollo': 					('Dev',	'K',	32,		[2, 11, 19, 27, 35]),
    'Diseño': 						('Dis',	'R',	36,		[2, 10, 19, 28, 38]),
    'PT': 							('PT',	'S',	40,		[2, 12, 22, 31, 40]),
    'QA': 							('QA',	'V',	60,		[2, 12, 23, 31, 41]),
    'Referentes Técnicos': 			('RT',	'L',	80,		[2, 9]),
    'Líderes':						('Lid',	'W',	44,		[2, 11, 21, 31, 41, 51, 61]),
    'Marketing':					('M',	'P',	48,		[2, 11, 21, 31, 41]),
    'Scrum Masters':				('SM',	'M',	68,		[2, 11, 21]),
    'People Care':					('PC',	'U',	52,		[2, 13, 24, 35, 44]),
    'Team Managers':				('TM',	'N',	64,		[2, 12])
}

# This is a matching between the auxiliar tabs and an array including the modes in which each tab is included
template_auxiliar_dictionary = {
    'Referencias matriz': 			[NEXT_EVALUATION, MANAGER_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION],
    'Feedback':				 		[NEXT_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION],
    'Satisfacción Laboral': 		[NEXT_EVALUATION, MANAGER_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION],
    'Desempeño Evaluadores': 		[MANAGER_EVALUATION],
    'Desempeño': 					[NEXT_EVALUATION, FIRST_EVALUATION],
    'Desempeño Intercambio':		[EXCHANGE_EVALUATION],
    'Objetivos y Capacitaciones': 	[NEXT_EVALUATION, MANAGER_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION],
    'Síntesis RID': 				[RID_EVALUATION],
    'Referencias': 					[NEXT_EVALUATION, AUTO_EVALUATION, MANAGER_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION, RID_EVALUATION],
}
