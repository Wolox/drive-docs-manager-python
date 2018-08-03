# -*- coding: utf-8 -*-

# Install pygsheets:
# pip3 install git+git://github.com/nithinmurali/pygsheets@staging
# Documentation for pygsheets: https://github.com/nithinmurali/pygsheets

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

# For file named: 'Template - Categorías 2.1'
# Use as key just the code shown in URL after https://docs.google.com/spreadsheets/d/
TEMPLATE_FILE_KEY = '10pGXiWVm6dzfwaSCXHWd0G4FetUgmIChCiq5Cm7dIKc'
# Timestamp for file. In case the found timestamp is different than this, execution must be aborted
TEMPLATE_TIMESTAMP = '2018-08-02T17:35:45.332Z'

# For file named: 'Rol Laboral - Evaluaciones de Desempeño (Respuestas)'
# Use as key just the code shown in URL after https://docs.google.com/spreadsheets/d/
ANSWERS_ROLE_FILE_KEY = '1felT_0RAVlG4FWFTbCkx3XMJjVSTd5sXqOLVYMzcRSo'

# Structs declarations

# This is a matching between a talent and a 4-tuple with:
# (talent_identifier, columnt for talent in answers_role_sheet, row for talent in 'Desempeño', row in tab for title talents):
struct_dictionary = {
	'Universales': 					('U',	'J',	24,		[2, 12, 23, 35, 45, 56, 67, 78, 88, 99]),
	'Administración y Finanzas': 	('AF',	'L',	48,		[2, 12, 22]),
	'Business Dev': 				('BD',	'M',	64,		[2, 12, 22, 32, 42, 52]),
	'Calidad': 						('C',	'P',	68,		[2, 12, 22, 32, 42]),
	'Desarrollo': 					('Dev',	'K',	28,		[2, 11, 19, 27, 35]),
	'Diseño': 						('Dis',	'N',	32,		[2, 10, 19, 28, 38]),
	'QA': 							('QA',	'O',	52,		[2, 12, 23, 31, 41]),
	'Referentes Técnicos': 			('RT',	'R',	72,		[2, 9]),
	'Líderes':						('Lid',	'U',	36,		[2, 11, 21, 31, 41, 51, 61]),
	'Marketing':					('M',	'V',	40,		[2, 11, 21, 31, 41]),
	'Scrum Masters':				('SM',	'S',	60,		[2, 11, 21]),
	'People Care':					('PC',	'Q',	44,		[2, 13, 24, 35, 44]),
	'Team Managers':				('TM',	'T',	56,		[2, 12])
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

# Open and return the template sheet
def get_template_sheet(google_credentials):
	template_sheet = google_credentials.open_by_key(TEMPLATE_FILE_KEY)
	print('')
	return template_sheet

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
				return destiny_sheet
		except:
			print('El valor ingresado \'' + destiny_file_key + '\' no es válido. Revisar y volver a intentar.')
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
				return answers_row_index
		except:
			print('El valor ingresado \'' + answers_row_index + '\' no es válido. Revisar y volver a intentar.')
			print('')

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

# If destiny sheet contains an evaluation for date to append, then the copy should not be re done
def copy_should_be_omitted(destiny_sheet, date_to_append):
	if destiny_sheet.sheet1.title.endswith(date_to_append):
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

# Copy tabs from template_sheet to destiny_sheet, letting the user select whether or not they are needed
def copy_tabs(destiny_sheet, template_sheet, date_to_append):
	destiny_worksheets = destiny_sheet.worksheets()
	template_worksheets = template_sheet.worksheets()

	# Worksheets in indexes 1 to 4 are read from template
	# Ignoring worksheet in tab 0, which is hidden: 'Referencias matriz (old)'
	worksheets_to_copy = [template_worksheets[1], template_worksheets[2], template_worksheets[3], template_worksheets[4]]

	# This is to discard non talent worksheets from template
	template_talent_worksheets = template_worksheets[5:len(template_worksheets)-1]

	# In case destiny contains multiple evaluations, only the last one must be duplicated, so a limit must be found
	# This limit is when the worksheet named 'Referencias' or 'Referencias N/YY' is found
	destiny_last_evaluation_worksheets = []
	if len(destiny_worksheets) > 1:
		limit_index = next(index for index, each in enumerate(destiny_worksheets) if each.index > 0 and each.title.startswith('Referencias')) + 1
		destiny_last_evaluation_worksheets = destiny_worksheets[0:limit_index]

	# Iterate over template talent worksheets copying each talent tab. In case the same tab already exists in destiny worksheet, then use it
	for i, worksheet in enumerate(template_talent_worksheets):
		# Once 'Referencias' or 'Referencias N/YY' is found, then break the cycle, to avoid keep on copying tabs from an older evaluation
		if worksheet.index > 0 and worksheet.title.startswith('Referencias'):
			break

		# Every worksheet is copied, in case it is already present in destiny_worksheet, it must be taken from there
		clean_title_to_search = worksheet.title if not worksheet.title[-1].isdigit() else worksheet.title[:-5]
		found_destiny_worksheet = list(filter(lambda each: each.title.startswith(clean_title_to_search), destiny_last_evaluation_worksheets))
		worksheet_to_copy = worksheet if not found_destiny_worksheet else found_destiny_worksheet[0]
		
		# In case 'Desarrollo' or 'Scrum Masters' is found in destiny_worksheet, it must be copied from template_worksheet since it may have updates
		if worksheet.title.startswith('Desarrollo') or worksheet.title.startswith('Scrum Masters'):
			worksheet_to_copy = worksheet

		worksheets_to_copy.append(worksheet_to_copy)

	# Worksheets in last index 'Referencias' or 'Referencias N/YY' is read from template
	worksheets_to_copy.append(template_worksheets[-1])

	print('')
	print('Comenzando copia de tabs al documento de evaluación...')
	print('')

	for i, worksheet in enumerate(worksheets_to_copy):
		# If title finishes in number it assumes the date needs to be removed and replaced by date_to_append
		title = (worksheet.title if not worksheet.title[-1].isdigit() else worksheet.title[:-5]) + ' ' + date_to_append
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

	print('Estado del documento actualizado!')
	print('')

	# For first time remove default worksheet, which after copying at the beggining, should be at the last index
	default_worksheet = list(filter(lambda each: each.title.startswith('Sheet'), destiny_sheet.worksheets()))
	if default_worksheet:
		print('Borrando tab default: ' + default_worksheet[0].title)
		destiny_sheet.del_worksheet(default_worksheet[0])
		print('Borrada tab default: ' + default_worksheet[0].title)
		print('')

	# Update cell with evaluation instance in 'Desempeño'
	instance_worksheet = destiny_sheet.worksheet_by_title('Desempeño' + ' ' + date_to_append)
	print('Actualizando instancia de evaluación en tab: ' + instance_worksheet.title)
	instance_worksheet.update_value('F17', date_to_append)
	print('Actualizada instancia de evaluación en tab: ' + instance_worksheet.title)
	print('')

	# Update talent cells 'Desarrollo' since it may have changes
	current_development_worksheet = destiny_sheet.worksheet_by_title('Desarrollo' + ' ' + date_to_append)
	previous_development_worksheet = list(filter(lambda each: each.title.startswith('Desarrollo') and each.index > current_development_worksheet.index, destiny_sheet.worksheets()))
	previous_development_worksheet = previous_development_worksheet[0] if previous_development_worksheet else None
	copy_talents_for_development(current_development_worksheet, previous_development_worksheet)

	# Update talent cells 'Scrum Masters' since it may have changes
	current_scrum_masters_worksheet = destiny_sheet.worksheet_by_title('Scrum Masters' + ' ' + date_to_append)
	previous_scrum_masters_worksheet = list(filter(lambda each: each.title.startswith('Scrum Masters') and each.index > current_scrum_masters_worksheet.index, destiny_sheet.worksheets()))
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

# Hide talents in destiny_sheet by reading the chosen ones from answers_role_sheet in answers_role_row
def hide_unused_talents(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append):
	print('')
	print('Ocultando talentos no seleccionados...')
	print('')

	for key, value in struct_dictionary.items():
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
	if not bool(answers_cell):
		# Get worksheet destiny based on name 'Desempeño' and date_to_append
		brief_worksheet = destiny_sheet.worksheet_by_title('Desempeño' + ' ' + date_to_append)
		cells_with_values_from = (title_row_in_brief + 2, 2)
		cells_with_values_to = (title_row_in_brief + 2, 1 + talents_amount)
		cells_with_values = brief_worksheet.get_values(start=cells_with_values_from, end=cells_with_values_to, returnas='cell')[0]
		# If there is no value different than '1' in the row, then the talent must be hidden
		if not True in list(map(lambda each: not each.value == '1', cells_with_values)):
			print('Ocultando fila en: ' + brief_worksheet.title)
			brief_worksheet.hide_rows(title_row_in_brief - 2, title_row_in_brief + 2)

	# In case there are no talents chosen from this worksheet, no hiding/showing of talents is necessary
	if not answers_cell:
		print('')
		return

	# Array with the positions of the cells from worksheet which contain the titles for the talents
	cells_with_titles_addresses = list(map(lambda each: 'B' + str(each), struct_dictionary[worksheet_name][3]))
	cells_with_titles = list(map(lambda each: worksheet_destiny.cell(each), cells_with_titles_addresses))

	# Array with the clean titles for the talents, after removing prefix
	worksheet_talents = list(map(lambda each: each.value[(len(each.value.split()[0]) + 1):], cells_with_titles))

	# Show every row before hiding to avoid hiding every row if there were already hidden ones
	worksheet_destiny.show_rows(0)

	# Hide every row, and then show only the selected ones
	worksheet_destiny.hide_rows(1, worksheet_destiny.rows - 1)

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
			worksheet_destiny.show_rows(show_from, show_to)
			print('Mostrando talento: ' + title)

	print('')

# Copy answers to destiny_sheet by reading them from answers_role_sheet in answers_role_row
def copy_answers_role(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append):
	# Get worksheet with name 'Respuestas de formulario 1' from answers_role_sheet
	answers_worksheet = answers_role_sheet.worksheet_by_title('Respuestas de formulario 1')

	# Get worksheet destiny based on name 'Satisfacción Laboral' and date_to_append
	destiny_worksheet = destiny_sheet.worksheet_by_title('Satisfacción Laboral' + ' ' + date_to_append)

	print('Realizando copia de respuestas del formulario de rol laboral a tab: ' + destiny_worksheet.title)
	print('')

	matching_dictionary = {
		# Question: "¿En qué grado te sentís satisfecho/a con tu rol laboral actual?"
		'D' + answers_role_row: ('B9', 'B10'),
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

# Script

google_credentials = get_google_credentials()
destiny_sheet = get_destiny_sheet(google_credentials)
template_sheet = get_template_sheet(google_credentials)
validate_updated_timestamp(template_sheet, TEMPLATE_TIMESTAMP)
date_to_append = get_date_to_append()
if not copy_should_be_omitted(destiny_sheet, date_to_append):
	copy_tabs(destiny_sheet, template_sheet, date_to_append)
	wait_for_quota_renewal()
answers_role_sheet = get_answers_role_sheet(google_credentials)
answers_role_row = get_answers_role_row(answers_role_sheet)
hide_unused_talents(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append)
copy_answers_role(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append)
on_finish()
