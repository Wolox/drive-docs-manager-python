import readchar
from constants import NEXT_EVALUATION, RID_EVALUATION, AUTO_EVALUATION, MANAGER_EVALUATION, \
	EXCHANGE_EVALUATION, FIRST_EVALUATION, UPDATE_FEEDBACK
from validations import is_valid_date_to_append

def get_date_to_append():
	"""Ask for date to append as evaluation instance"""
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

def get_title():
	"""Ask for title to create the document or folder"""
	while True:
		title = input('Ingrese el título:')
		if not title:
			print('El valor ingresado \'' + title + '\' no es válido. Revisar y volver a intentar.')
			continue

		print('Título ingresado: \'' + title + '\'')
		print('Es correcto? Presione \'s/n\'.')
		if readchar.readchar() == 's':
			print('Título confirmado...')
			print('')
			return title

def get_mode():
	"""Ask for mode to run the script"""
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

def get_rid_instance_to_append():
	"""Ask for RID instance to append"""
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