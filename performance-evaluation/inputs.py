import readchar

from constants import NEXT_EVALUATION, RID_EVALUATION, AUTO_EVALUATION, MANAGER_EVALUATION, \
	EXCHANGE_EVALUATION, FIRST_EVALUATION, UPDATE_FEEDBACK, TEMPLATE_AUXILIAR_FILE_KEY, \
	TEMPLATE_TALENT_FILE_KEY, OPERATION_COPY_TABS, OPERATION_BUILD_EVALUATION_FORM, \
	OPERATION_HIDE_TALENTS, OPERATION_COPY_ANSWERS, OPERATION_COPY_FEEDBACK, \
	OPERATION_CREATE_SHEET, operations_by_mode_dictionary, folders_dictionary, template_talents_dictionary, \
	template_auxiliar_dictionary
from google_sheets import set_google_credentials, create_sheet, get_sheet_by_key, \
	get_feedback_sheet, get_answers_role_sheet, get_answers_role_row
from google_drive import get_key_by_path
from operations import copy_tabs, build_evaluation_form, hide_unused_talents, copy_answers_role, copy_feedback
from validations import is_valid_date_to_append

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