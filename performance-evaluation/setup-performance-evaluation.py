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
TEMPLATE_TIMESTAMP = '2018-07-10T20:32:59.555Z'

# For file named: 'Rol Laboral - Evaluaciones de Desempeño (Respuestas)'
# Use as key just the code shown in URL after https://docs.google.com/spreadsheets/d/
ANSWERS_ROLE_FILE_KEY = '1felT_0RAVlG4FWFTbCkx3XMJjVSTd5sXqOLVYMzcRSo'

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
		destiny_sheet = google_credentials.open_by_key(destiny_file_key)
		print('Se va a actualizar el documento con el nombre: \'' + destiny_sheet.title + '\'')
		print('Es correcto? Presione \'s/n\'.')
		if readchar.readchar() == 's':
			print('Abriendo documento para comenzar a trabajar...')
			return destiny_sheet

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
		print('Se van a copiar las respuestas para: \'' + answers_role_sheet.sheet1.cell('C' + answers_row_index).value + '\'')
		print('Es correcto? Presione \'s/n\'.')
		if readchar.readchar() == 's':
			print('Leyendo respuestas...')
			return answers_row_index

def get_date_to_append():
	while True:
		date_to_append = input('Ingresar instancia de evaluación. Ejemplo: \'1/18\': ')
		print('Instancia ingresada: \'' + date_to_append + '\'')
		print('Es correcto? Presione \'s/n\'.')
		if readchar.readchar() == 's':
			print('Instancia confirmada...')
			return date_to_append

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
def copy_tabs(destiny_sheet, template_sheet):
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
		
		# In case 'Desarrollo' is found in destiny_worksheet, it must be copied from template_worksheet anyway since it may have updates
		if worksheet.title.startswith('Desarrollo'):
			worksheet_to_copy = worksheet

		worksheets_to_copy.append(worksheet_to_copy)

	# Worksheets in last index 'Referencias' or 'Referencias N/YY' is read from template
	worksheets_to_copy.append(template_worksheets[-1])

	# Asking for date to append replacing the date in template
	date_to_append = get_date_to_append()
	print('')

	print('Comenzando copia de tabs al documento de evaluación...')
	print('')

	for i, worksheet in enumerate(worksheets_to_copy):
		# If title finishes in number it assumes the date needs to be removed and replaced by date_to_append
		title = (worksheet.title if not worksheet.title[-1].isdigit() else worksheet.title[:-5]) + ' ' + date_to_append
		print('Copiando tab (' + str(i) + '): ' + title)
		new_worksheet = destiny_sheet.add_worksheet(title, src_worksheet=worksheet)
		new_worksheet.index = i

	print('')
	print('Copia de tabs finalizada!')
	
	print('')
	print('Actualizando estado del documento...')

	# Updating index for every previously existing worksheet in destiny_sheet to avoid duplicated indexes after copying
	for i, worksheet in enumerate(destiny_sheet.worksheets()):
		if not worksheet.title.endswith(date_to_append):
			worksheet.index = worksheet.index + len(worksheets_to_copy)

	# For first time remove default worksheet, which after copying at the beggining, should be at the last index
	default_worksheet = list(filter(lambda each: each.title.startswith('Sheet'), destiny_sheet.worksheets()))
	if default_worksheet:
		print('Borrando tab default: ' + default_worksheet[0].title)
		destiny_sheet.del_worksheet(default_worksheet[0])

	# Update cell with evaluation instance in 'Desempeño'
	instance_worksheet = list(filter(lambda each: each.title == 'Desempeño ' + date_to_append, destiny_sheet.worksheets()))[0]
	print('Actualizando instancia de evaluación en tab: ' + instance_worksheet.title)
	instance_worksheet.update_value('F17', date_to_append)

	# Update talent cells 'Desarrollo' since it may have changes
	updated_destiny_worksheets = destiny_sheet.worksheets()
	current_development_worksheet = list(filter(lambda each: each.title == 'Desarrollo ' + date_to_append, updated_destiny_worksheets))[0]
	print('Actualizando tab: ' + current_development_worksheet.title)
	previous_development_worksheet = list(filter(lambda each: each.title.startswith('Desarrollo') and each.index > current_development_worksheet.index, updated_destiny_worksheets))
	previous_development_worksheet = previous_development_worksheet[0] if previous_development_worksheet else None
	copy_talents_for_development(current_development_worksheet, previous_development_worksheet)

	print('')

def copy_talents_for_development(current_worksheet, previous_worksheet):
	# If previous_worksheet is None, then this talent is evaluated by first time, so nowhere to copy from
	if not previous_worksheet:
		return

	# If 'B12' contains 'Dev2', then previous_worksheet has the older format, otherwise the matching is direct
	need_to_adapt = previous_worksheet.cell('B12').value.startswith('Dev2')

	columns_to_update = ['G', 'H', 'I']
	rows_to_update = [15, 16, 23, 24, 31, 32, 39, 40]

	# For talents which did not change, just copy directly
	for column in columns_to_update:
		for row in rows_to_update:
			cell_to_read = previous_worksheet.cell(column + str(row if not need_to_adapt else row + 1))
			current_worksheet.update_value(column + str(row), cell_to_read.value)

	# For 'Dev1', give special treatment
	if not need_to_adapt:
		rows_to_update = [6, 7, 8]
		for column in columns_to_update:
			for row in rows_to_update:
				cell_to_read = previous_worksheet.cell(column + str(row))
				current_worksheet.update_value(column + str(row), cell_to_read.value)
	else:
		# Cells for 'Aprendizaje' -> 'Aprendizaje'
		matching_dictionary = {'G6': 'G6', 'H6': 'H6', 'I6': 'I6'}
		for key, value in matching_dictionary.items():
			cell_from = previous_worksheet.cell(key)
			current_worksheet.update_value(value, cell_from.value)

		# Cells for 'Calidad de código' -> 'Calidad y transmisión'
		matching_dictionary = {'G9': 'G8', 'H9': 'H8', 'I9': 'I8'}
		for key, value in matching_dictionary.items():
			cell_from = previous_worksheet.cell(key)
			cell_to = current_worksheet.cell(value)
			if len(cell_to.value) == 0:
				current_worksheet.update_value(value, cell_from.value)

# Hide talents in destiny_sheet by reading the chosen ones from answers_role_sheet in answers_role_row
def hide_unused_talents(destiny_sheet, answers_role_sheet, answers_role_row):
	print('Ocultando talentos no seleccionados...')
	print('')

	matching_dictionary = {
		'Universales': ('U', 'J'),
		'Administración y Finanzas': ('AF', 'L'),
		'Business Dev': ('BD', 'M'),
		'Calidad': ('C', 'P'),
		'Desarrollo': ('Dev', 'K'),
		'Diseño': ('Dis', 'N'),
		'QA': ('QA', 'O'),
		'Referentes Técnicos': ('RT', 'R'),
		'Líderes': ('Lid', 'U'),
		'Marketing': ('M', 'V'),
		'Scrum Masters': ('SM', 'S'),
		'People Care': ('PC', 'Q'),
		'Team Managers': ('TM', 'T')
	}

	for key, value in matching_dictionary.items():
		hide_unused_talents_in_single_worksheet(destiny_sheet, key, value[0], answers_role_sheet, answers_role_row, value[1])

	print('Ocultamiento de talentos finalizado!')
	print('')

def hide_unused_talents_in_single_worksheet(destiny_sheet, worksheet_name, worksheet_identifier, answers_role_sheet, answers_role_row, answers_role_column):
	print('Comenzando a mostrar talentos para tab: ' + worksheet_name)

	worksheet_destiny_array = destiny_sheet.worksheets()
	worksheet_destiny_index = next(index for index, each in enumerate(worksheet_destiny_array) if each.title.startswith(worksheet_name))
	worksheet_destiny = worksheet_destiny_array[worksheet_destiny_index]

	# Cell from answers document with the selected talents. Info is separated by comma
	answers_cell = answers_role_sheet.sheet1.cell(answers_role_column + answers_role_row).value.lower()

	# Hide worksheet if there are no talents chosen from it
	worksheet_destiny.hidden = not bool(answers_cell)
	print('(visible)' if bool(answers_cell) else '(oculto)')

	# In case there are no talents chosen from this worksheet, no hiding/showing of talents is necessary
	if not answers_cell:
		print('')
		return

	# Array with the positions of the cells from worksheet which contain the titles for the talents
	# cells_with_titles = worksheet_destiny.find('u[0-9][0-9]?', searchByRegex=True)
	cells_with_titles = provisional_find_for_worksheet(worksheet_destiny, worksheet_name)

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

# This function is provisional and will be removed once the 'Insufficient tokens for quota' error is fixed
def provisional_find_for_worksheet(worksheet, title):
	matching_dictionary = {
		'Universales': ['B2', 'B12', 'B23', 'B35', 'B45', 'B56', 'B67', 'B78', 'B88', 'B99'],
		'Administración y Finanzas': ['B2', 'B12', 'B22'],
		'Business Dev': ['B2', 'B12', 'B22', 'B32', 'B42', 'B52'],
		'Calidad': ['B2', 'B12', 'B22', 'B32', 'B42'],
		'Desarrollo': ['B2', 'B11', 'B19', 'B27', 'B35'],
		'Diseño': ['B2', 'B10', 'B19', 'B28', 'B38'],
		'QA': ['B2', 'B12', 'B23', 'B31', 'B41'],
		'Referentes Técnicos': ['B2', 'B9'],
		'Líderes': ['B2', 'B11', 'B21', 'B31', 'B41', 'B51', 'B61'],
		'Marketing': ['B2', 'B11', 'B21', 'B31', 'B41'],
		'Scrum Masters': ['B2', 'B12', 'B23', 'B33', 'B43'],
		'People Care': ['B2', 'B13', 'B24', 'B35', 'B44'],
		'Team Managers': ['B2', 'B12']
	}

	for key, value in matching_dictionary.items():
		if title == key:
			return list(map(lambda each: worksheet.cell(each), value))

# Copy answers to destiny_sheet by reading them from answers_role_sheet in answers_role_row
def copy_answers_role(destiny_sheet, answers_role_sheet, answers_role_row):
	# Worksheet in index 0 to copy the answers 'Respuestas de formulario 1'
	answers_worksheet_array = answers_role_sheet.worksheets()
	answers_worksheet_index = next(index for index, each in enumerate(answers_worksheet_array) if each.title.startswith('Respuestas de formulario 1'))
	answers_worksheet = answers_role_sheet.worksheets()[answers_worksheet_index]

	# Worksheet in index 1 to paste the answers 'Satisfacción Laboral'
	destiny_worksheet = list(filter(lambda each: each.index == 1, destiny_sheet.worksheets()))[0]

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
copy_tabs(destiny_sheet, template_sheet)
wait_for_quota_renewal()
answers_role_sheet = get_answers_role_sheet(google_credentials)
answers_role_row = get_answers_role_row(answers_role_sheet)
hide_unused_talents(destiny_sheet, answers_role_sheet, answers_role_row)
copy_answers_role(destiny_sheet, answers_role_sheet, answers_role_row)
on_finish()
