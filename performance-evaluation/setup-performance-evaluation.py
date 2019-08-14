# -*- coding: utf-8 -*-

import pygsheets
import readchar
import time

from constants import *
from google_sheets import validate_updated_timestamp, get_google_credentials, get_template_sheets, \
	get_destiny_sheet, get_auto_evaluation_sheet, get_manager_evaluation_sheet, \
	get_exchange_evaluation_sheet, get_feedback_sheet, get_answers_role_sheet
from operations import copy_tabs, build_evaluation_form, hide_unused_talents, copy_answers_role, copy_feedback

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
