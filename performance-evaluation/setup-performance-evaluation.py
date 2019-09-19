# -*- coding: utf-8 -*-
import readchar
import time

from constants import NEXT_EVALUATION, RID_EVALUATION, AUTO_EVALUATION, MANAGER_EVALUATION, \
	EXCHANGE_EVALUATION, FIRST_EVALUATION, UPDATE_FEEDBACK, TEMPLATE_AUXILIAR_FILE_KEY, \
	TEMPLATE_TALENT_FILE_KEY, OPERATION_COPY_TABS, OPERATION_BUILD_EVALUATION_FORM, \
	OPERATION_HIDE_TALENTS, OPERATION_COPY_ANSWERS, OPERATION_COPY_FEEDBACK, \
	operations_by_mode_dictionary, folders_dictionary, template_talents_dictionary, \
	template_auxiliar_dictionary
from google_sheets import set_google_credentials, create_sheet, get_sheet_by_key, \
	get_feedback_sheet, get_answers_role_sheet, get_answers_role_row
from google_drive import get_key_by_path
from operations import copy_tabs, build_evaluation_form, hide_unused_talents, copy_answers_role, copy_feedback

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

# Ask for title to create the document
def get_title():
	while True:
		title = input('Ingresar el título del nuevo documento. Ejemplo: \'Salas, Julián - Formulario de evaluación\': ')
		if not title:
			print('El valor ingresado \'' + title + '\' no es válido. Revisar y volver a intentar.')
			continue

		print('Título ingresado: \'' + title + '\'')
		print('Es correcto? Presione \'s/n\'.')
		if readchar.readchar() == 's':
			print('Título confirmado...')
			print('')
			return title

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
set_google_credentials()
template_auxiliar_sheet = get_sheet_by_key(TEMPLATE_AUXILIAR_FILE_KEY)
template_talent_sheet = get_sheet_by_key(TEMPLATE_TALENT_FILE_KEY)

print('A continuación te vamos a pedir los siguientes datos:')
print(' - Instancia de la evaluación')
print(' - Título del documento. En caso de no generar ninguno, no se utilizará')
print('')
date_to_append = get_date_to_append()
import ipdb; ipdb.set_trace()
title = get_title()
mode = get_mode()

mode_build_evaluation_form = mode in operations_by_mode_dictionary[OPERATION_BUILD_EVALUATION_FORM]
mode_hide_talents = mode in operations_by_mode_dictionary[OPERATION_HIDE_TALENTS]
mode_copy_answers = mode in operations_by_mode_dictionary[OPERATION_COPY_ANSWERS]

if mode_build_evaluation_form or mode_hide_talents or mode_copy_answers:
	answers_role_sheet = get_answers_role_sheet()
	answers_role_row = get_answers_role_row(answers_role_sheet)

if not folders_dictionary[mode]['need_child_key']:
	if not folders_dictionary[mode]['parent_key']:
		key = get_key_by_path(folders_dictionary[mode]['parent_path'])
	else:
		key = folders_dictionary[mode]['parent_key']
	destiny_sheet = create_sheet(title, key)
else:
	parent_key = get_key_by_path(folders_dictionary[mode]['parent_path'])
	key = get_key_by_path(parent_key)
	destiny_sheet = get_sheet_by_key(key)

rid_instance_to_append = get_rid_instance_to_append() if mode == RID_EVALUATION else None

if mode in operations_by_mode_dictionary[OPERATION_COPY_TABS]:
	if not copy_should_be_omitted(destiny_sheet, date_to_append):
		copy_tabs(destiny_sheet, template_auxiliar_sheet, template_talent_sheet, date_to_append, rid_instance_to_append, mode)
		if not mode == RID_EVALUATION:
			wait_for_quota_renewal()

if mode in operations_by_mode_dictionary[OPERATION_BUILD_EVALUATION_FORM]:
	auto_evaluation_sheet = None
	manager_evaluation_sheet = None
	exchange_evaluation_sheet = None
	if mode == EXCHANGE_EVALUATION:
		auto_evaluation_key = get_key_by_path(folders_dictionary[AUTO_EVALUATION]['parent_key'])
		auto_evaluation_sheet = get_sheet_by_key(auto_evaluation_key)
		manager_evaluation_key = get_key_by_path(folders_dictionary[MANAGER_EVALUATION]['parent_key'])
		manager_evaluation_sheet = get_sheet_by_key(manager_evaluation_key)
	if mode == FIRST_EVALUATION:
		exchange_evaluation_key = get_key_by_path(folders_dictionary[EXCHANGE_EVALUATION]['parent_key'])
		exchange_evaluation_sheet = get_sheet_by_key(exchange_evaluation_key)
	build_evaluation_form(destiny_sheet, date_to_append, mode, auto_evaluation_sheet, manager_evaluation_sheet, exchange_evaluation_sheet, answers_role_sheet, answers_role_row)
	wait_for_quota_renewal()

if mode in operations_by_mode_dictionary[OPERATION_HIDE_TALENTS]:
	hide_unused_talents(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append, mode)

if mode in operations_by_mode_dictionary[OPERATION_COPY_ANSWERS]:
	copy_answers_role(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append, mode)

if mode in operations_by_mode_dictionary[OPERATION_COPY_FEEDBACK]:
	feedback_sheet = get_feedback_sheet()
	copy_feedback(destiny_sheet, feedback_sheet, date_to_append)

check_specific_talents(destiny_sheet, answers_role_sheet, answers_role_row, mode) if mode != RID_EVALUATION else None

on_finish()
