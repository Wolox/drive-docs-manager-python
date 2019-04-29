# -*- coding: utf-8 -*-

# Install pygsheets:
# pip3 install git+git://github.com/nithinmurali/pygsheets@75594dc58a5a9671edea2283369bb190aac36fb3

# Install readchar:
# pip3 install readchar
# Documentation for readchar: https://github.com/magmax/python-readchar

# For error "ImportError: No module named oauth2client.file"
# Then do: pip3 install --upgrade oauth2client

# Follow this to create credentials:
# https://pygsheets.readthedocs.io/en/latest/authorizing.html

# Quick access to my provisional credentials:
# https://console.developers.google.com/apis/dashboard?folder=&organizationId=22904326237&project=peoplecare-automation

import pygsheets
import readchar
import time

# Constant declarations

# Path to credentials file. Read documentation https://pygsheets.readthedocs.io/en/latest/authorizing.html to create new credentials
CREDENTIALS_FILE_PATH = 'client_secret.json'

# For file named: 'Template - Evaluación Auxiliares'
TEMPLATE_AUXILIAR_FILE_KEY = '1eB-6j0xc9qeFVtuF3xXfajiWD9e_TcFdxxTHHdlAjyU'
TEMPLATE_AUXILIAR_FILE_TIMESTAMP = '2018-09-07T21:26:15.870Z'

# For file named: 'Template - Evaluación Talentos'
TEMPLATE_TALENT_FILE_KEY = '1D04Q-IQ67F1wTAgk3oEr1Q902DJd4ulv666mBlcar6c'
TEMPLATE_TALENT_FILE_TIMESTAMP = '2018-09-07T21:26:28.775Z'

# For file named: 'Rol Laboral - Evaluaciones de Desempeño (Respuestas)'
ANSWERS_ROLE_FILE_KEY = '1felT_0RAVlG4FWFTbCkx3XMJjVSTd5sXqOLVYMzcRSo'

# Mode for running script for ever evaluation after the first one.
NEXT_EVALUATION = "NEXT_EVALUATION"
# Mode for creating the auto evaluation form.
AUTO_EVALUATION = "AUTO_EVALUATION"
# Mode for creating the manager evaluation form.
MANAGER_EVALUATION = "MANAGER_EVALUATION"
# Mode for creating the exchage evaluation form.
EXCHANGE_EVALUATION = "EXCHANGE_EVALUATION"
# Mode for creating the first evaluation form after filling the agreement form.
FIRST_EVALUATION = "FIRST_EVALUATION"
# Mode for creating a RID evaluation.
RID_EVALUATION = "RID_EVALUATION"
# Mode for updating feedback for a given evaluation.
UPDATE_FEEDBACK = "UPDATE_FEEDBACK"

OPERATION_COPY_TABS = "OPERATION_COPY_TABS"
OPERATION_BUILD_EVALUATION_FORM = "OPERATION_BUILD_EVALUATION_FORM"
OPERATION_HIDE_TALENTS = "OPERATION_HIDE_TALENTS"
OPERATION_COPY_ANSWERS = "OPERATION_COPY_ANSWERS"
OPERATION_COPY_FEEDBACK = "OPERATION_COPY_FEEDBACK"

operations_by_mode_dictionary = {
	OPERATION_COPY_TABS: 				[NEXT_EVALUATION, AUTO_EVALUATION, MANAGER_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION, RID_EVALUATION],
	OPERATION_BUILD_EVALUATION_FORM:	[EXCHANGE_EVALUATION, FIRST_EVALUATION],
	OPERATION_HIDE_TALENTS: 			[NEXT_EVALUATION, AUTO_EVALUATION, MANAGER_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION],
	OPERATION_COPY_ANSWERS: 			[NEXT_EVALUATION, AUTO_EVALUATION, MANAGER_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION],
	OPERATION_COPY_FEEDBACK:			[NEXT_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION, UPDATE_FEEDBACK],
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

# Function declarations

# Authenticate using OAuth from a given credentials file path and return credentials
def get_google_credentials():
	print('Autenticando...')
	google_credentials = pygsheets.authorize(CREDENTIALS_FILE_PATH)
	print('Autenticación exitosa!')
	print('')
	return google_credentials

# Checks the last updated timestamp por template. In case it was modified, script code may have to be adapted
def validate_updated_timestamp(spreadsheet, timestamp):
	if spreadsheet.updated != timestamp:
		print('Fecha de última modificación: ' + spreadsheet.updated)
		print('Cancelando script. Archivo \'' + spreadsheet.title + '\' modificado, revisar script.')
		print('')
		exit()

# Open and return the template sheets
def get_template_sheets(google_credentials):
	template_auxiliar_sheet = google_credentials.open_by_key(TEMPLATE_AUXILIAR_FILE_KEY)
	# validate_updated_timestamp(template_auxiliar_sheet, TEMPLATE_AUXILIAR_FILE_TIMESTAMP)
	template_talent_sheet = google_credentials.open_by_key(TEMPLATE_TALENT_FILE_KEY)
	# validate_updated_timestamp(template_talent_sheet, TEMPLATE_TALENT_FILE_TIMESTAMP)
	return (template_auxiliar_sheet, template_talent_sheet)

# Ask for the document to update and return the destiny sheet
def get_destiny_sheet(google_credentials):
	while True:
		destiny_file_key = input('Ingresar la key del documento a actualizar: ')
		try:
			destiny_sheet = google_credentials.open_by_key(destiny_file_key)
			print('Se va a actualizar el documento con el nombre: \'' + destiny_sheet.title + '\'')
			print('Es correcto? Presione \'s/n\'.')
			if readchar.readchar() == 's':
				print('Abriendo documento para comenzar a trabajar...')
				print('')
				return destiny_sheet
		except:
			print('El valor ingresado \'' + destiny_file_key + '\' no es válido. Revisar y volver a intentar.')
			print('')

# Ask for the document to read the auto evaluation
def get_auto_evaluation_sheet(google_credentials):
	while True:
		destiny_file_key = input('Ingresar la key del documento de auto evaluación: ')
		try:
			destiny_sheet = google_credentials.open_by_key(destiny_file_key)
			print('Se va a utilizar el documento de auto evaluación con el nombre: \'' + destiny_sheet.title + '\'')
			print('Es correcto? Presione \'s/n\'.')
			if readchar.readchar() == 's':
				print('Abriendo documento para comenzar a trabajar...')
				print('')
				return destiny_sheet
		except:
			print('El valor ingresado \'' + destiny_file_key + '\' no es válido. Revisar y volver a intentar.')
			print('')

# Ask for the document to read the evaluation
def get_manager_evaluation_sheet(google_credentials):
	while True:
		destiny_file_key = input('Ingresar la key del documento de evaluación: ')
		try:
			destiny_sheet = google_credentials.open_by_key(destiny_file_key)
			print('Se va a utilizar el documento de evaluación con el nombre: \'' + destiny_sheet.title + '\'')
			print('Es correcto? Presione \'s/n\'.')
			if readchar.readchar() == 's':
				print('Abriendo documento para comenzar a trabajar...')
				print('')
				return destiny_sheet
		except:
			print('El valor ingresado \'' + destiny_file_key + '\' no es válido. Revisar y volver a intentar.')
			print('')

# Ask for the document to read the exchange evaluation
def get_exchange_evaluation_sheet(google_credentials):
	while True:
		destiny_file_key = input('Ingresar la key del documento de evaluación de intercambio: ')
		try:
			destiny_sheet = google_credentials.open_by_key(destiny_file_key)
			print('Se va a utilizar el documento de evaluación de intercambio con el nombre: \'' + destiny_sheet.title + '\'')
			print('Es correcto? Presione \'s/n\'.')
			if readchar.readchar() == 's':
				print('Abriendo documento para comenzar a trabajar...')
				print('')
				return destiny_sheet
		except:
			print('El valor ingresado \'' + destiny_file_key + '\' no es válido. Revisar y volver a intentar.')
			print('')

# Ask for the document to read the feedback
def get_feedback_sheet(google_credentials):
	while True:
		feedback_file_key = input('Ingresar la key del documento de feedback: ')
		try:
			feedback_sheet = google_credentials.open_by_key(feedback_file_key)
			print('Se va a utilizar el documento de feedback con el nombre: \'' + feedback_sheet.title + '\'')
			print('Es correcto? Presione \'s/n\'.')
			if readchar.readchar() == 's':
				print('Abriendo documento para comenzar a trabajar...')
				print('')
				return feedback_sheet
		except:
			print('El valor ingresado \'' + feedback_file_key + '\' no es válido. Revisar y volver a intentar.')
			print('')

# Open and return the answers role sheet
def get_answers_role_sheet(google_credentials):
	print('Abriendo documento de respuestas de formulario para comenzar la copia...')
	answers_role_sheet = google_credentials.open_by_key(ANSWERS_ROLE_FILE_KEY)
	print('')
	return answers_role_sheet

# Ask for the row number in answers role sheet
def get_answers_role_row(answers_role_sheet):
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

# Ask for RID instance to append
def get_rid_instance_to_append():
	while True:
		rid_instance_to_append = input('Ingresar instancia de RID. Ejemplo: \'1\': ')
		if not rid_instance_to_append.isnumeric() or int(rid_instance_to_append) < 1 or int(rid_instance_to_append) > 9:
			print('El valor ingresado \'' + rid_instance_to_append + '\' no es válido. Revisar y volver a intentar.')
			continue

		print('Instancia ingresada: \'' + rid_instance_to_append + '\'')
		print('Es correcto? Presione \'s/n\'.')
		if readchar.readchar() == 's':
			print('Instancia confirmada...')
			print('')
			return rid_instance_to_append

# Ask for date to append as evaluation instance
def get_date_to_append():
	while True:
		date_to_append = input('Ingresar instancia de evaluación. Ejemplo: \'1/18\': ')
		if not is_valid_date_to_append(date_to_append):
			print('El valor ingresado \'' + date_to_append + '\' no es válido. Revisar y volver a intentar.')
			continue

		print('Instancia ingresada: \'' + date_to_append + '\'')
		print('Es correcto? Presione \'s/n\'.')
		if readchar.readchar() == 's':
			print('Instancia confirmada...')
			print('')
			return date_to_append

def is_valid_date_to_append(date_to_append):
	if not len(date_to_append) == 4:
		return False
	if not date_to_append[:1].isnumeric():
		return False
	if not date_to_append[1:2] == '/':
		return False
	if not date_to_append[3:].isnumeric():
		return False
	return True

# Ask for mode to run the script
def get_mode():
	mode_next_evaluation = '1- Preparar informe de desempeño individual'
	mode_auto_evaluation = '2- Preparar formulario de autoevaluación'
	mode_manager_evaluation = '3- Preparar formulario de evaluación para evaluador'
	mode_exchange_evaluation = '4- Preparar formulario de evaluación para intercambio'
	mode_first_evaluation = '5- Preparar informe de desempeño individual por primera vez'
	mode_rid_evaluation = '6- Preparar informe de RID'
	mode_update_feedback = '7- Actualizar feedback de una evaluación existente'
	while True:
		print('Ingresar modalidad de ejecución:')
		print(mode_next_evaluation)
		print(mode_auto_evaluation)
		print(mode_manager_evaluation)
		print(mode_exchange_evaluation)
		print(mode_first_evaluation)
		print(mode_rid_evaluation)
		print(mode_update_feedback)

		input_mode_number = input('Ingresar opción \'1-7\': ')
		mode = None
		mode = NEXT_EVALUATION if input_mode_number == '1' else mode
		mode = AUTO_EVALUATION if input_mode_number == '2' else mode
		mode = MANAGER_EVALUATION if input_mode_number == '3' else mode
		mode = EXCHANGE_EVALUATION if input_mode_number == '4' else mode
		mode = FIRST_EVALUATION if input_mode_number == '5' else mode
		mode = RID_EVALUATION if input_mode_number == '6' else mode
		mode = UPDATE_FEEDBACK if input_mode_number == '7' else mode

		if mode is None:
			print('El valor ingresado \'' + input_mode_number + '\' no es válido. Revisar y volver a intentar.')
			print('')
			continue

		input_mode = None
		input_mode = mode_next_evaluation if mode == NEXT_EVALUATION else input_mode
		input_mode = mode_auto_evaluation if mode == AUTO_EVALUATION else input_mode
		input_mode = mode_manager_evaluation if mode == MANAGER_EVALUATION else input_mode
		input_mode = mode_exchange_evaluation if mode == EXCHANGE_EVALUATION else input_mode
		input_mode = mode_first_evaluation if mode == FIRST_EVALUATION else input_mode
		input_mode = mode_rid_evaluation if mode == RID_EVALUATION else input_mode
		input_mode = mode_update_feedback if mode == UPDATE_FEEDBACK else input_mode

		print('Modalidad ingresada: \'' + input_mode[3:] + '\'')
		print('Es correcto? Presione \'s/n\'.')
		if readchar.readchar() == 's':
			print('Modalidad confirmada...')
			print('')
			return mode

# If destiny sheet contains an evaluation for date to append, then the copy should not be re done
def copy_should_be_omitted(destiny_sheet, date_to_append):
	already_present = next((index for index, each in enumerate(destiny_sheet.worksheets()) if not 'RID' in each.title and each.title.endswith(date_to_append)), None)
	if already_present:
		print('')
		print('Ya existe una evaluación para la instancia: \'' + date_to_append + '\'')
		print('La copia de tabs no se va a realizar. Sólo se van a ocultar las tabs y talentos, y copiar las respuestas.')
		print('Continuar? Presione \'s/n\'.')
		if not readchar.readchar() == 's':
			print('Abortando ejecución. No se realizaron cambios en el documento.')
			exit()
		print('')
		return True
	return False

def wait_for_quota_renewal():
	remaining_seconds = 120
	waiting_seconds = 15
	while remaining_seconds > 0:
		print('Por favor esperar ' + str(remaining_seconds) + ' segundos...')
		time.sleep(waiting_seconds)
		remaining_seconds = remaining_seconds - waiting_seconds
	print('Continuando ejecución...')
	print('')

def on_finish():
	print('Ejecución finalizada con éxito! No olvidar verificar el estado del documento de forma manual.')

# Copy tabs from templates to destiny_sheet
def copy_tabs(destiny_sheet, template_auxiliar_sheet, template_talent_sheet, date_to_append, rid_instance_to_append, mode):
	destiny_worksheets = destiny_sheet.worksheets()
	template_auxiliar_worksheets = template_auxiliar_sheet.worksheets()
	template_talent_worksheets = template_talent_sheet.worksheets()

	# Auxiliar worksheets to be copied at the beginning, excluding 'Referencias' in last index, to be copied at the ending
	worksheets_to_copy = []
	for each in template_auxiliar_worksheets[:-1]:
		if mode in template_auxiliar_dictionary[each.title]:
			worksheets_to_copy.append(each)

	# In case destiny contains multiple evaluations, only the last one must be duplicated, so a limit must be found
	# This limit is when the worksheet named 'Referencias' or 'Referencias N/YY' is found
	destiny_last_evaluation_worksheets = []
	if mode == NEXT_EVALUATION:
		start_index = next(index for index, each in enumerate(destiny_worksheets) if not 'RID' in each.title and 'Referencias matriz' in each.title)
		end_index = next(index for index, each in enumerate(destiny_worksheets) if not 'RID' in each.title and 'Referencias' in each.title and not 'Referencias matriz' in each.title) + 1
		destiny_last_evaluation_worksheets = destiny_worksheets[start_index:end_index]
	if mode == RID_EVALUATION and len(destiny_worksheets) > 1:
		end_index = next(index for index, each in enumerate(destiny_worksheets) if each.index > 0 and 'Referencias' in each.title and not 'Referencias matriz' in each.title) + 1
		destiny_last_evaluation_worksheets = destiny_worksheets[0:end_index]

	# Iterate over template talent worksheets copying each talent tab. In case the same tab already exists in destiny worksheet, then use it
	for i, worksheet in enumerate(template_talent_worksheets):
		# Once 'Referencias' or 'Referencias N/YY' is found, then break the cycle, to avoid keep on copying tabs from an older evaluation
		if worksheet.index > 0 and worksheet.title.startswith('Referencias'):
			break

		# If RID is contained in title, we know for sure, it also ends in digit.
		if 'RID' in worksheet.title:
			clean_title_to_search = worksheet.title[:-11]
		# If RID is not contained, it may end in digit.
		elif worksheet.title[-1].isdigit():
			clean_title_to_search = worksheet.title[:-5]
		# If RID is not contained and it does not end in digit.
		else:
			clean_title_to_search = worksheet.title

		# Every worksheet is copied, in case it is already present in destiny_worksheet, it must be taken from there
		found_destiny_worksheet = list(filter(lambda each: each.title.startswith(clean_title_to_search), destiny_last_evaluation_worksheets))
		worksheet_to_copy = worksheet if not found_destiny_worksheet else found_destiny_worksheet[0]

		# In case 'Desarrollo' or 'Scrum Masters' is found in destiny_worksheet, it must be copied from template_worksheet since it may have updates
		if worksheet.title.startswith('Desarrollo') or worksheet.title.startswith('Scrum Masters'):
			worksheet_to_copy = worksheet

		worksheets_to_copy.append(worksheet_to_copy)

	# Worksheet in last index 'Referencias'
	worksheets_to_copy.append(template_auxiliar_worksheets[-1])

	print('Comenzando copia de tabs al documento de evaluación...')
	print('')

	for i, worksheet in enumerate(worksheets_to_copy):

		# Need to check this tab in a different clause, to avoid entering in second if.
		if 'Síntesis RID' in worksheet.title:
			clean_title = worksheet.title
		# If RID is contained in title, we know for sure, it also ends in digit.
		elif 'RID' in worksheet.title:
			clean_title = worksheet.title[:-11]
		# If RID is not contained, it may end in digit.
		elif worksheet.title[-1].isdigit():
			clean_title = worksheet.title[:-5]
		# If RID is not contained and it does not end in digit.
		else:
			clean_title = worksheet.title

		title = clean_title if not mode == RID_EVALUATION else clean_title + ' ' + 'RID' + ' ' + rid_instance_to_append
		title = title + ' ' + date_to_append
		print('Copiando tab (' + str(i) + '): ' + title + ' -- desde: ' + worksheet.title)
		new_worksheet = destiny_sheet.add_worksheet(title, src_worksheet=worksheet)
		new_worksheet.index = i

	print('')
	print('Copia de tabs finalizada!')
	print('')

	print('Actualizando estado del documento.')

	# Updating index for every previously existing worksheet in destiny_sheet to avoid duplicated indexes after copying
	for i, worksheet in enumerate(destiny_sheet.worksheets()):
		if not worksheet.title.endswith(date_to_append):
			worksheet.index = worksheet.index + len(worksheets_to_copy)

	# Removing unused cols for every talent worksheet in case the worksheet was copied from template
	for key, value in template_talents_dictionary.items():
		worksheet_title = key if not mode == RID_EVALUATION else key + ' ' + 'RID' + ' ' + rid_instance_to_append
		worksheet_title = worksheet_title + ' ' + date_to_append
		worksheet = destiny_sheet.worksheet_by_title(worksheet_title)
		if worksheet.cols == 19:
			column_to_remove_start_index = 16 if mode == EXCHANGE_EVALUATION else 7
			column_to_remove_amount = 3 if mode == EXCHANGE_EVALUATION else 9
			worksheet.delete_cols(column_to_remove_start_index, number=column_to_remove_amount)

	print('Estado del documento actualizado!')
	print('')

	# For new document evaluations, the number of copied sheets more the default one will match the current ones
	if len(destiny_sheet.worksheets()) == len(worksheets_to_copy) + 1:
		default_worksheet = list(filter(lambda each: not each.title.endswith(date_to_append), destiny_sheet.worksheets()))[0]
		print('Borrando tab default: ' + default_worksheet.title)
		destiny_sheet.del_worksheet(default_worksheet)
		print('Borrada tab default: ' + default_worksheet.title)
		print('')

	# Update cell with evaluation instance in 'Desempeño'
	# No need to check for RID since this tab is not present in this mode.
	if mode in template_auxiliar_dictionary['Desempeño']:
		instance_worksheet = destiny_sheet.worksheet_by_title('Desempeño' + ' ' + date_to_append)
		print('Actualizando instancia de evaluación en tab: ' + instance_worksheet.title)
		instance_worksheet.update_value('F21', date_to_append)
		print('Actualizada instancia de evaluación en tab: ' + instance_worksheet.title)
		print('')

	# Update cell with evaluation instance in 'Desempeño Evaluadores'
	# No need to check for RID since this tab is not present in this mode.
	if mode in template_auxiliar_dictionary['Desempeño Evaluadores']:
		instance_worksheet = destiny_sheet.worksheet_by_title('Desempeño Evaluadores' + ' ' + date_to_append)
		print('Actualizando instancia de evaluación en tab: ' + instance_worksheet.title)
		instance_worksheet.update_value('F21', date_to_append)
		print('Actualizada instancia de evaluación en tab: ' + instance_worksheet.title)
		print('')

	# Update cell with evaluation instance in 'Desempeño Intercambio'
	# No need to check for RID since this tab is not present in this mode.
	if mode in template_auxiliar_dictionary['Desempeño Intercambio']:
		instance_worksheet = destiny_sheet.worksheet_by_title('Desempeño Intercambio' + ' ' + date_to_append)
		print('Actualizando instancia de evaluación en tab: ' + instance_worksheet.title)
		instance_worksheet.update_value('F21', date_to_append)
		print('Actualizada instancia de evaluación en tab: ' + instance_worksheet.title)
		print('')

	# Update talent cells 'Desarrollo' since it may have changes
	# Current one may be RID or not. Previous one has to be last evaluation but not RID
	current_development_worksheet_title = 'Desarrollo' if not mode == RID_EVALUATION else 'Desarrollo' + ' ' + 'RID' + ' ' + rid_instance_to_append
	current_development_worksheet_title = current_development_worksheet_title + ' ' + date_to_append
	current_development_worksheet = destiny_sheet.worksheet_by_title(current_development_worksheet_title)
	previous_development_worksheet = list(filter(lambda each: each.title.startswith('Desarrollo') and not 'RID' in each.title and each.index > current_development_worksheet.index, destiny_sheet.worksheets()))
	previous_development_worksheet = previous_development_worksheet[0] if previous_development_worksheet else None
	copy_talents_for_development(current_development_worksheet, previous_development_worksheet)

	# Update talent cells 'Scrum Masters' since it may have changes
	# Current one may be RID or not. Previous one has to be last evaluation but not RID
	current_scrum_masters_worksheet_title = 'Scrum Masters' if not mode == RID_EVALUATION else 'Scrum Masters' + ' ' + 'RID' + ' ' + rid_instance_to_append
	current_scrum_masters_worksheet_title = current_scrum_masters_worksheet_title + ' ' + date_to_append
	current_scrum_masters_worksheet = destiny_sheet.worksheet_by_title(current_scrum_masters_worksheet_title)
	previous_scrum_masters_worksheet = list(filter(lambda each: each.title.startswith('Scrum Masters') and not 'RID' in each.title and each.index > current_scrum_masters_worksheet.index, destiny_sheet.worksheets()))
	previous_scrum_masters_worksheet = previous_scrum_masters_worksheet[0] if previous_scrum_masters_worksheet else None
	copy_talents_for_scrum_masters(current_scrum_masters_worksheet, previous_scrum_masters_worksheet)

def copy_talents_for_development(current_worksheet, previous_worksheet):
	print('Actualizando tab: ' + current_worksheet.title)

	# If previous_worksheet is None, then this talent is evaluated by first time, so nowhere to copy from
	if not previous_worksheet:
		print('Nada para copiar...')
		print('')
		return

	# If 'B12' contains 'Dev2', then previous_worksheet has the older format, otherwise the matching is direct
	need_to_adapt = previous_worksheet.cell('B12').value.startswith('Dev2')

	cells_if_not_need_to_adapt = {
		('G6', 'I8'): ('G6', 'I8'),		# Dev1
		('G15', 'I16'): ('G15', 'I16'), # Dev2
		('G23', 'I24'): ('G23', 'I24'), # Dev3
		('G31', 'I32'): ('G31', 'I32'), # Dev4
		('G39', 'I40'): ('G39', 'I40'), # Dev5
	}
	cells_if_need_to_adapt = {
		('G6', 'I6'): ('G6', 'I6'), 	# Dev1: 'Aprendizaje'
		('G9', 'I9'): ('G8', 'I8'),		# Dev1: 'Calidad de código'
		('G16', 'I17'): ('G15', 'I16'), # Dev2
		('G24', 'I25'): ('G23', 'I24'), # Dev3
		('G32', 'I33'): ('G31', 'I32'), # Dev4
		('G40', 'I41'): ('G39', 'I40'), # Dev5
	}
	ranges_to_update = cells_if_not_need_to_adapt if not need_to_adapt else cells_if_need_to_adapt
	for key, value in ranges_to_update.items():
		talents = previous_worksheet.get_values(start=key[0], end=key[1], returnas='matrix')
		current_worksheet.update_values(crange=value[0] + ':' + value[1], values=talents)

	print('Actualizada tab: ' + current_worksheet.title)
	print('')

def copy_talents_for_scrum_masters(current_worksheet, previous_worksheet):
	print('Actualizando tab: ' + current_worksheet.title)

	# If previous_worksheet is None, then this talent is evaluated by first time, so nowhere to copy from
	if not previous_worksheet:
		print('Nada para copiar...')
		print('')
		return

	# If 'B12' contains 'SM2', then previous_worksheet has the older format, otherwise the matching is direct
	need_to_adapt = previous_worksheet.cell('B12').value.startswith('SM2')

	cells_if_not_need_to_adapt = {
		('G6', 'I8'): ('G6', 'I8'),		# SM1
		('G15', 'I18'): ('G15', 'I18'), # SM2
		('G25', 'I28'): ('G25', 'I28'), # SM3
	}
	cells_if_need_to_adapt = {
		('G6', 'I6'): ('G6', 'I6'), 	# SM1: 'Metodología y agilidad del proceso de desarrollo'
		('G8', 'I8'): ('G7', 'I7'),		# SM1: 'Mejora Continua'
		('G49', 'I49'): ('G8', 'I8'), 	# SM1: 'Mecanismos de seguimiento y control'
		('G16', 'I19'): ('G15', 'I18'), # SM2
		('G28', 'I30'): ('G25', 'I27'), # SM3
		('G39', 'I39'): ('G28', 'I28'), # SM3: 'Acciones en relación a generación de comportamientos de participación y generación de ideas.'
	}
	ranges_to_update = cells_if_not_need_to_adapt if not need_to_adapt else cells_if_need_to_adapt
	for key, value in ranges_to_update.items():
		talents = previous_worksheet.get_values(start=key[0], end=key[1], returnas='matrix')
		current_worksheet.update_values(crange=value[0] + ':' + value[1], values=talents)

	print('Actualizada tab: ' + current_worksheet.title)
	print('')

# Copy tabs from templates to destiny_sheet
def copy_feedback(destiny_sheet, feedback_sheet, date_to_append):
	print('Copiando feedback...')
	print('')

	# Get last feedback worksheet, in case there is any, to get last feedback copied
	feedback_limit_index = 0
	found_last_feedback_worksheet = list(filter(lambda each: each.title.startswith('Feedback') and not each.title.endswith(date_to_append) and bool(each.cell('A7').value), destiny_sheet.worksheets()))
	if found_last_feedback_worksheet:
		feedback_rows = found_last_feedback_worksheet[0].get_all_values()
		feedback_rows = feedback_rows[6:]
		feedback_rows = list(filter(lambda each: len(each[0]) > 0, feedback_rows))
		last_feedback_timestamp = feedback_rows[-1][0] if feedback_rows else None
		if last_feedback_timestamp:
			feedback_limit_index = next(index for index, each in enumerate(feedback_rows) if each[0] == last_feedback_timestamp) + 1

	# Get last feedback used, in order to avoid copying twice the same row
	worksheet_feedback = feedback_sheet.worksheets()[0]
	feedback_rows = worksheet_feedback.get_all_values()
	feedback_rows = feedback_rows[1:]
	feedback_rows = list(filter(lambda each: len(each[0]) > 0, feedback_rows))
	feedback_rows_to_copy = feedback_rows[feedback_limit_index:]

	if len(feedback_rows_to_copy) > 0:
		feedback_rows_to_copy = list(map(lambda each: each[:6], feedback_rows_to_copy))

		worksheet_destiny = destiny_sheet.worksheet_by_title('Feedback' + ' ' + date_to_append)
		feedback_range = 'A7:F' + str(7 + len(feedback_rows_to_copy))
		worksheet_destiny.update_values(crange=feedback_range, values=feedback_rows_to_copy)

	# Hiding feedback sheet in case there is no feedback for this evaluation
	worksheet_destiny = destiny_sheet.worksheet_by_title('Feedback' + ' ' + date_to_append)
	worksheet_destiny.hidden = len(feedback_rows_to_copy) == 0
	print('Tab de: ' + worksheet_destiny.title)
	print('(oculta)' if len(feedback_rows_to_copy) == 0 else '(visible)')

	print('Copiado feedback de ' + str(len(feedback_rows_to_copy)) + ' woloxers!')
	print('')

def build_evaluation_form(destiny_sheet, auto_evaluation_sheet, manager_evaluation_sheet, exchange_evaluation_sheet, answers_role_sheet, answers_role_row, date_to_append, mode):
	print('Copiando talentos seleccionados...')
	print('')

	for key, value in template_talents_dictionary.items():
		build_evaluation_form_in_single_worksheet(destiny_sheet, auto_evaluation_sheet, manager_evaluation_sheet, exchange_evaluation_sheet, key, value[0], answers_role_sheet, answers_role_row, value[1], date_to_append, len(value[3]), mode)

	print('Copia de talentos finalizada!')
	print('')

# Copies evaluated talents from a sheet to destiny.
# In case mode is EXCHANGE_EVALUATION it uses auto_evaluation_sheet and manager_evaluation_sheet.
# In case mode is FIRST_EVALUATION it uses exchange_evaluation_sheet.
def build_evaluation_form_in_single_worksheet(destiny_sheet, auto_evaluation_sheet, manager_evaluation_sheet, exchange_evaluation_sheet, worksheet_name, worksheet_identifier, answers_role_sheet, answers_role_row, answers_role_column, date_to_append, talents_amount, mode):
	# Get worksheet destiny based on worksheet_name and date_to_append
	worksheet_destiny = destiny_sheet.worksheet_by_title(worksheet_name + ' ' + date_to_append)

	# Get auto evaluation worksheet based on worksheet_name and date_to_append
	try:
		auto_evaluation_worksheet = auto_evaluation_sheet.worksheet_by_title(worksheet_name + ' ' + date_to_append) if mode == EXCHANGE_EVALUATION else None
	except pygsheets.exceptions.WorksheetNotFound:
		print('Ausente tab: ' + worksheet_destiny.title + '. Ignorando copia de talentos.')
		print('')
		return

	# Get manager evaluation worksheet based on worksheet_name and date_to_append
	try:
		manager_evaluation_worksheet = manager_evaluation_sheet.worksheet_by_title(worksheet_name + ' ' + date_to_append) if mode == EXCHANGE_EVALUATION else None
	except pygsheets.exceptions.WorksheetNotFound:
		print('Ausente tab: ' + worksheet_destiny.title + '. Ignorando copia de talentos.')
		print('')
		return

	# Get exchange evaluation worksheet based on worksheet_name and date_to_append
	try:
		exchange_evaluation_worksheet = exchange_evaluation_sheet.worksheet_by_title(worksheet_name + ' ' + date_to_append) if mode == FIRST_EVALUATION else None
	except pygsheets.exceptions.WorksheetNotFound:
		print('Ausente tab: ' + worksheet_destiny.title + '. Ignorando copia de talentos.')
		print('')
		return

	# Cell from answers document with the selected talents. Info is separated by comma
	answers_cell = answers_role_sheet.sheet1.cell(answers_role_column + answers_role_row).value.lower()

	print('Comenzando a copiar talentos para tab: ' + worksheet_destiny.title)
	print('(copiando)' if bool(answers_cell) else '(nada para copiar)')

	# Array with the positions of the cells from worksheet which contain the titles for the talents
	cells_with_titles_addresses = list(map(lambda each: 'B' + str(each), template_talents_dictionary[worksheet_name][3]))
	cells_with_titles = list(map(lambda each: worksheet_destiny.cell(each), cells_with_titles_addresses))

	# Array with the clean titles for the talents, after removing prefix
	worksheet_talents = list(map(lambda each: each.value[(len(each.value.split()[0]) + 1):], cells_with_titles))

	# Iterate over each of the titles that may appear in answers_role_sheet first sheet, copying those that are present in answers
	for i, title in enumerate(worksheet_talents):
		# If title is present in anwers, then the talent must he shown
		should_copy_by_match = title.lower() in answers_cell
		# This really makes me sad, but there are some talents that don't match between their names and the names in the answers form, so I ended up doing this
		should_copy_by_exception_1 = worksheet_identifier == 'C' and title.lower() == 'calidad y mejora continua (calidad)' and 'calidad y mejora continua' in answers_cell
		should_copy_by_exception_2 = worksheet_identifier == 'Dis' and title.lower() == 'comunicacion eficaz (diseño)' and 'comunicación eficaz' in answers_cell
		should_copy_by_exception_3 = worksheet_identifier == 'Dis' and title.lower() == 'resolución de problemas' and 'resolución de probelmas' in answers_cell
		should_copy_by_exception_4 = worksheet_identifier == 'SM' and title.lower() == 'colaboración / facilitador' and 'colaboración y facilitador' in answers_cell
		should_copy_by_exception_5 = worksheet_identifier == 'SM' and title.lower() == 'iniciativa / autonomía' and 'iniciativa y autonomía' in answers_cell
		if should_copy_by_match or should_copy_by_exception_1 or should_copy_by_exception_2 or should_copy_by_exception_3 or should_copy_by_exception_4:
			copy_from = list(filter(lambda each: each.value.startswith(worksheet_identifier + str(i + 1)), cells_with_titles))[0].row + 3
			copy_to = worksheet_destiny.rows if i + 1 == len(worksheet_talents) else list(filter(lambda each: each.value.startswith(worksheet_identifier + str(i + 2)), cells_with_titles))[0].row - 3

			if mode == EXCHANGE_EVALUATION:
				start_range_copy = 'G' + str(copy_from)
				end_range_copy = 'I' + str(copy_to)

				range_auto_evaluation = 'G' + str(copy_from) + ':' + 'I' + str(copy_to)
				talent_from_auto_evaluation = auto_evaluation_worksheet.get_values(start=start_range_copy, end=end_range_copy, majdim='COLUMNS')
				worksheet_destiny.update_values(crange=range_auto_evaluation, values=talent_from_auto_evaluation, majordim='COLUMNS')

				range_manager_evaluation = 'J' + str(copy_from) + ':' + 'L' + str(copy_to)
				talent_from_manager_evaluation = manager_evaluation_worksheet.get_values(start=start_range_copy, end=end_range_copy, majdim='COLUMNS')
				worksheet_destiny.update_values(crange=range_manager_evaluation, values=talent_from_manager_evaluation, majordim='COLUMNS')

			if mode == FIRST_EVALUATION:
				start_range_copy = 'M' + str(copy_from)
				end_range_copy = 'O' + str(copy_to)

				range_exchange_evaluation = 'G' + str(copy_from) + ':' + 'I' + str(copy_to)
				talent_from_exchange_evaluation = exchange_evaluation_worksheet.get_values(start=start_range_copy, end=end_range_copy, majdim='COLUMNS')
				worksheet_destiny.update_values(crange=range_exchange_evaluation, values=talent_from_exchange_evaluation, majordim='COLUMNS')

			print('Copiando talento: ' + title)

	print('')

# Hide talents in destiny_sheet by reading the chosen ones from answers_role_sheet in answers_role_row
def hide_unused_talents(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append):
	print('Ocultando talentos no seleccionados...')
	print('')

	for key, value in template_talents_dictionary.items():
		hide_unused_talents_in_single_worksheet(destiny_sheet, key, value[0], answers_role_sheet, answers_role_row, value[1], date_to_append, value[2], len(value[3]))

	print('Ocultamiento de talentos finalizado!')
	print('')

def hide_unused_talents_in_single_worksheet(destiny_sheet, worksheet_name, worksheet_identifier, answers_role_sheet, answers_role_row, answers_role_column, date_to_append, title_row_in_brief, talents_amount):
	# Get worksheet destiny based on worksheet_name and date_to_append
	worksheet_destiny = destiny_sheet.worksheet_by_title(worksheet_name + ' ' + date_to_append)

	# Cell from answers document with the selected talents. Info is separated by comma
	answers_cell = answers_role_sheet.sheet1.cell(answers_role_column + answers_role_row).value.lower()

	# Hide worksheet if there are no talents chosen from it
	worksheet_destiny.hidden = not bool(answers_cell)
	print('Comenzando a mostrar talentos para tab: ' + worksheet_destiny.title)
	print('(visible)' if bool(answers_cell) else '(oculto)')

	# If no talents were chosen, the row in brief may have to be hidden if no talent was ever evaluated
	if mode in template_auxiliar_dictionary['Desempeño'] and not bool(answers_cell):
		# Get worksheet destiny based on name 'Desempeño' and date_to_append
		brief_worksheet = destiny_sheet.worksheet_by_title('Desempeño' + ' ' + date_to_append)
		cells_with_values_from = (title_row_in_brief + 2, 2)
		cells_with_values_to = (title_row_in_brief + 2, 1 + talents_amount)
		cells_with_values = brief_worksheet.get_values(start=cells_with_values_from, end=cells_with_values_to, returnas='cell')[0]
		# If there is no value different than '1' in the row, then the talent must be hidden
		if not True in list(map(lambda each: not each.value == '1', cells_with_values)):
			print('Ocultando fila en: ' + brief_worksheet.title)
			brief_worksheet.hide_dimensions(title_row_in_brief - 2, title_row_in_brief + 2, dimension="ROWS")

	# In case there are no talents chosen from this worksheet, no hiding/showing of talents is necessary
	if not answers_cell:
		print('')
		return

	# Array with the positions of the cells from worksheet which contain the titles for the talents
	cells_with_titles_addresses = list(map(lambda each: 'B' + str(each), template_talents_dictionary[worksheet_name][3]))
	cells_with_titles = list(map(lambda each: worksheet_destiny.cell(each), cells_with_titles_addresses))

	# Array with the clean titles for the talents, after removing prefix
	worksheet_talents = list(map(lambda each: each.value[(len(each.value.split()[0]) + 1):], cells_with_titles))

	# Show every row before hiding to avoid hiding every row if there were already hidden ones
	worksheet_destiny.show_dimensions(0, dimension="ROWS")

	# Hide every row, and then show only the selected ones
	worksheet_destiny.hide_dimensions(1, worksheet_destiny.rows - 1, dimension="ROWS")

	# Iterate over each of the titles that may appear in answers_role_sheet first sheet, showing from destiny those that are present in answers
	for i, title in enumerate(worksheet_talents):
		# If title is present in anwers, then the talent must he shown
		should_show_by_match = title.lower() in answers_cell
		# This really makes me sad, but there are some talents that don't match between their names and the names in the answers form, so I ended up doing this
		should_show_by_exception_1 = worksheet_identifier == 'C' and title.lower() == 'calidad y mejora continua (calidad)' and 'calidad y mejora continua' in answers_cell
		should_show_by_exception_2 = worksheet_identifier == 'Dis' and title.lower() == 'comunicacion eficaz (diseño)' and 'comunicación eficaz' in answers_cell
		should_show_by_exception_3 = worksheet_identifier == 'Dis' and title.lower() == 'resolución de problemas' and 'resolución de probelmas' in answers_cell
		should_show_by_exception_4 = worksheet_identifier == 'SM' and title.lower() == 'colaboración / facilitador' and 'colaboración y facilitador' in answers_cell
		should_show_by_exception_5 = worksheet_identifier == 'SM' and title.lower() == 'iniciativa / autonomía' and 'iniciativa y autonomía' in answers_cell
		if should_show_by_match or should_show_by_exception_1 or should_show_by_exception_2 or should_show_by_exception_3 or should_show_by_exception_4:
			show_from = list(filter(lambda each: each.value.startswith(worksheet_identifier + str(i + 1)), cells_with_titles))[0].row - 1
			show_to = worksheet_destiny.rows if i + 1 == len(worksheet_talents) else list(filter(lambda each: each.value.startswith(worksheet_identifier + str(i + 2)), cells_with_titles))[0].row - 1
			worksheet_destiny.show_dimensions(show_from, show_to, dimension="ROWS")
			print('Mostrando talento: ' + title)

	print('')

# Copy answers to destiny_sheet by reading them from answers_role_sheet in answers_role_row
def copy_answers_role(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append, mode):
	# If this tab is not used for this mode, then answers must not be copied
	if mode not in template_auxiliar_dictionary['Satisfacción Laboral']:
		print('Omitiendo copia de respuestas del formulario de rol laboral')
		print('')
		return

	# Answers sheet from 'Rol Laboral'.
	answers_worksheet = answers_role_sheet.sheet1

	# Get worksheet destiny based on name 'Satisfacción Laboral' and date_to_append
	destiny_worksheet = destiny_sheet.worksheet_by_title('Satisfacción Laboral' + ' ' + date_to_append)

	print('Realizando copia de respuestas del formulario de rol laboral a tab: ' + destiny_worksheet.title)
	print('')

	matching_dictionary = {
		# Question: "¿En qué grado te sentís satisfecho/a con tu rol laboral actual?"
		'E' + answers_role_row: ('B9', 'B10'),
		# Question: "¿Por qué lo calificarías así?"
		'F' + answers_role_row: ('B12', 'B13'),
		# Question: ¿Qué otros aprendizajes te gustaría poder realizar en el transcurso de los próximos 6 meses?"
		'G' + answers_role_row: ('B15', 'B16'),
		# Question: ¿Qué otro rol y/o en qué otra área te gustaría poder trabajar a futuro dentro de Wolox?"
		'H' + answers_role_row: ('B18', 'B19'),
		# Question: "¿Qué aprendizajes considerás necesarios poder adquirir a fin de ejercer dicho rol? ¿En qué te gustaría capacitarte?"
		'I' + answers_role_row: ('B21', 'B22')
	}

	for key, value in matching_dictionary.items():
		destiny_question_cell = value[0]
		destiny_answer_cell = value[1]
		print('Copiando respuesta a: \'' + destiny_worksheet.cell(destiny_question_cell).value + '\'')
		answer = answers_worksheet.cell(key).value
		destiny_worksheet.update_value(destiny_answer_cell, answer)

	print('')
	print('Copia de respuestas finalizada!')
	print('')

def check_specific_talents(destiny_sheet, answers_role_sheet, answers_role_row, mode):

	# Update checks in 'Desempeño' tab in order to read the proper specific talent.
	# No need to check for RID since this tab is not present in this mode.
	performance_tabs_to_update = ['Desempeño Evaluadores', 'Desempeño', 'Desempeño Intercambio']

	for each in performance_tabs_to_update:
		if mode in template_auxiliar_dictionary[each]:
			instance_worksheet = destiny_sheet.worksheet_by_title(each + ' ' + date_to_append)

			# Update cell 'Impacto' check, in order to set it to TRUE in case last evaluation uses it.
			leaders_worksheet = destiny_sheet.worksheet_by_title('Líderes' + ' ' + date_to_append)
			impact_cell_has_value = bool(leaders_worksheet.cell('G55').value.strip())
			print('Actualizando check de Impacto en tab: ' + instance_worksheet.title)
			instance_worksheet.update_value('J27', 'TRUE' if impact_cell_has_value else 'FALSE')
			print('Actualizada check de Impacto en tab: ' + instance_worksheet.title)
			print('')

			development_talents_chosen_cell = template_talents_dictionary['Desarrollo'][1] + answers_role_row
			development_talents_chosen = bool(answers_role_sheet.sheet1.cell(development_talents_chosen_cell).value.strip())
			print('Actualizando check de Desarrollo en tab: ' + instance_worksheet.title)
			instance_worksheet.update_value('G27', 'TRUE' if development_talents_chosen else 'FALSE')
			print('Actualizada check de Desarrollo en tab: ' + instance_worksheet.title)
			print('')

			design_talents_chosen_cell = template_talents_dictionary['Diseño'][1] + answers_role_row
			design_talents_chosen = bool(answers_role_sheet.sheet1.cell(design_talents_chosen_cell).value.strip())
			print('Actualizando check de Diseño en tab: ' + instance_worksheet.title)
			instance_worksheet.update_value('H27', 'TRUE' if design_talents_chosen else 'FALSE')
			print('Actualizada check de Diseño en tab: ' + instance_worksheet.title)
			print('')

			product_thinking_talents_chosen_cell = template_talents_dictionary['PT'][1] + answers_role_row
			product_thinking_talents_chosen = bool(answers_role_sheet.sheet1.cell(product_thinking_talents_chosen_cell).value.strip())
			print('Actualizando check de PT en tab: ' + instance_worksheet.title)
			instance_worksheet.update_value('I27', 'TRUE' if product_thinking_talents_chosen else 'FALSE')
			print('Actualizada check de PT en tab: ' + instance_worksheet.title)
			print('')

# Script

google_credentials = get_google_credentials()
(template_auxiliar_sheet, template_talent_sheet) = get_template_sheets(google_credentials)
mode = get_mode()
destiny_sheet = get_destiny_sheet(google_credentials)
date_to_append = get_date_to_append()
rid_instance_to_append = get_rid_instance_to_append() if mode == RID_EVALUATION else None

if mode in operations_by_mode_dictionary[OPERATION_COPY_TABS]:
	if not copy_should_be_omitted(destiny_sheet, date_to_append):
		copy_tabs(destiny_sheet, template_auxiliar_sheet, template_talent_sheet, date_to_append, rid_instance_to_append, mode)
		if not mode == RID_EVALUATION:
			wait_for_quota_renewal()

mode_build_evaluation_form = mode in operations_by_mode_dictionary[OPERATION_BUILD_EVALUATION_FORM]
mode_hide_talents = mode in operations_by_mode_dictionary[OPERATION_HIDE_TALENTS]
mode_copy_answers = mode in operations_by_mode_dictionary[OPERATION_COPY_ANSWERS]

if mode_build_evaluation_form or mode_hide_talents or mode_copy_answers:
	answers_role_sheet = get_answers_role_sheet(google_credentials)
	answers_role_row = get_answers_role_row(answers_role_sheet)

if mode in operations_by_mode_dictionary[OPERATION_BUILD_EVALUATION_FORM]:
	auto_evaluation_sheet = get_auto_evaluation_sheet(google_credentials) if mode == EXCHANGE_EVALUATION else None
	manager_evaluation_sheet = get_manager_evaluation_sheet(google_credentials) if mode == EXCHANGE_EVALUATION else None
	exchange_evaluation_sheet = get_exchange_evaluation_sheet(google_credentials) if mode == FIRST_EVALUATION else None
	build_evaluation_form(destiny_sheet, auto_evaluation_sheet, manager_evaluation_sheet, exchange_evaluation_sheet, answers_role_sheet, answers_role_row, date_to_append, mode)
	wait_for_quota_renewal()

if mode in operations_by_mode_dictionary[OPERATION_HIDE_TALENTS]:
	hide_unused_talents(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append)

if mode in operations_by_mode_dictionary[OPERATION_COPY_ANSWERS]:
	copy_answers_role(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append, mode)

if mode in operations_by_mode_dictionary[OPERATION_COPY_FEEDBACK]:
	feedback_sheet = get_feedback_sheet(google_credentials)
	copy_feedback(destiny_sheet, feedback_sheet, date_to_append)

check_specific_talents(destiny_sheet, answers_role_sheet, answers_role_row, mode)

on_finish()
