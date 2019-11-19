# -*- coding: utf-8 -*-
import pygsheets
from constants import NEXT_EVALUATION, RID_EVALUATION, EXCHANGE_EVALUATION, FIRST_EVALUATION, \
    template_talents_dictionary, operations_by_mode_dictionary, template_auxiliar_dictionary

def copy_tabs(destiny_sheet, template_auxiliar_sheet, template_talent_sheet, date_to_append, rid_instance_to_append, mode):
	"""Copy tabs from templates to destiny_sheet"""
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

		# In case 'Desarrollo' or 'Scrum Masters' or 'Diseño' is found in destiny_worksheet, it must be copied from template_worksheet since it may have updates
		if worksheet.title.startswith('Desarrollo') or worksheet.title.startswith('Scrum Masters') or worksheet.title.startswith('Diseño'):
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
	for key, _ in template_talents_dictionary.items():
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

	# Update talent cells 'Diseño' since it may have changes
	# Current one may be RID or not. Previous one has to be last evaluation but not RID
	current_design_worksheet_title = 'Diseño' if not mode == RID_EVALUATION else 'Diseño' + ' ' + 'RID' + ' ' + rid_instance_to_append
	current_design_worksheet_title = current_design_worksheet_title + ' ' + date_to_append
	current_design_worksheet = destiny_sheet.worksheet_by_title(current_design_worksheet_title)
	previous_design_worksheet = list(filter(lambda each: each.title.startswith('Diseño') and not 'RID' in each.title and each.index > current_design_worksheet.index, destiny_sheet.worksheets()))
	previous_design_worksheet = previous_design_worksheet[0] if previous_design_worksheet else None
	copy_talents_for_design(current_design_worksheet, previous_design_worksheet)

def build_evaluation_form(destiny_sheet, date_to_append, mode, auto_evaluation_sheet, manager_evaluation_sheet, exchange_evaluation_sheet, answers_role_sheet, answers_role_row):
	"""this function call _build_evaluation_form_in_single_worksheet"""
	print('Copiando talentos seleccionados...')
	print('')

	for key, value in template_talents_dictionary.items():
		_build_evaluation_form_in_single_worksheet(destiny_sheet, auto_evaluation_sheet, manager_evaluation_sheet, exchange_evaluation_sheet, key, value[0], answers_role_sheet, answers_role_row, value[1], date_to_append, len(value[3]), mode)

	print('Copia de talentos finalizada!')
	print('')

def hide_unused_talents(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append, mode):
	"""Hide talents in destiny_sheet by reading the chosen ones from answers_role_sheet in answers_role_row"""
	print('Ocultando talentos no seleccionados...')
	print('')

	for key, value in template_talents_dictionary.items():
		_hide_unused_talents_in_single_worksheet(destiny_sheet, key, value[0], answers_role_sheet, answers_role_row, value[1], date_to_append, value[2], len(value[3]), mode)

	print('Ocultamiento de talentos finalizado!')
	print('')

def copy_answers_role(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append, mode):
	"""Copy answers to destiny_sheet by reading them from answers_role_sheet in answers_role_row"""
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

def copy_feedback(destiny_sheet, feedback_sheet, date_to_append):
	"""Copy the feedback to the evaluation form"""
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

def copy_talents_for_development(current_worksheet, previous_worksheet):
	"""specific copy for Development talents"""
	print('Actualizando tab: ' + current_worksheet.title)

	# If previous_worksheet is None, then this talent is evaluated by first time, so nowhere to copy from
	if not previous_worksheet:
		print('Nada para copiar...')
		print('')
		return

	# If 'B19' ends with 'Testing' instead '(Deprecado)', then previous_worksheet has the older format, otherwise the matching is direct
	need_to_adapt = previous_worksheet.cell('B19').value.endswith('Testing')

	cells_if_not_need_to_adapt = {
		('G6', 'I8'): ('G6', 'I8'),		# Dev1
		('G15', 'I16'): ('G15', 'I16'), # Dev2
		('G23', 'I24'): ('G23', 'I24'), # Dev3 (Deprecado)
		('G31', 'I35'): ('G31', 'I35'), # Dev4 Desarrollo Front-End & Mobile
		('G42', 'I45'): ('G42', 'I45'), # Dev5 Desarrollo Backend
		('G52', 'I54'): ('G52', 'I54'), # Dev6 Infraestructura/Arquitectura
		('G61', 'I63'): ('G61', 'I63'), # Dev7 Comunidad de desarrollo
	}
	cells_if_need_to_adapt = {
		('G6', 'I8'): ('G6', 'I8'),		# Dev1
		('G15', 'I16'): ('G15', 'I16'), # Dev2
		('G23', 'I24'): ('G23', 'I24'), # Dev3 (Deprecado)
		('G31', 'I32'): ('G31', 'I32'), # Dev4
		('G39', 'I39'): ('G42', 'I42'), # Dev5: Integración con clientes (front, mobile, externos
	}
	ranges_to_update = cells_if_not_need_to_adapt if not need_to_adapt else cells_if_need_to_adapt
	for key, value in ranges_to_update.items():
		talents = previous_worksheet.get_values(start=key[0], end=key[1], returnas='matrix')
		current_worksheet.update_values(crange=value[0] + ':' + value[1], values=talents)

	print('Actualizada tab: ' + current_worksheet.title)
	print('')

def copy_talents_for_scrum_masters(current_worksheet, previous_worksheet):
	"""specific copy for SM talents"""
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

def copy_talents_for_design(current_worksheet, previous_worksheet):
	"""specific copy for Design talents"""
	print('Actualizando tab: ' + current_worksheet.title)

	# If previous_worksheet is None, then this talent is evaluated by first time, so nowhere to copy from
	if not previous_worksheet:
		print('Nada para copiar...')
		print('')
		return

	# If row_number 48 exists and  value ends with 'Motion Design', then previous_worksheet has the older format, otherwise the matching is direct
	all_cells = previous_worksheet.get_all_values(include_empty_rows=True, include_tailing_empty=True, returnas='cells')
	last_row = all_cells[-1][-1]
	if last_row.row >= 48:
		need_to_adapt = False if previous_worksheet.cell('B48').value.endswith('Motion Design') else True
	else:
		need_to_adapt = True

	cells_if_not_need_to_adapt = {
		('G6', 'I7'): ('G6', 'I7'),		# Dis1
		('G14', 'I16'): ('G14', 'I16'), # Dis2
		('G23', 'I25'): ('G23', 'I25'), # Dis3
		('G32', 'I35'): ('G32', 'I35'), # Dis4
		('G42', 'I45'): ('G42', 'I45'), # Dis5
		('G52', 'I53'): ('G52', 'I53'), # Dis6
	}
	cells_if_need_to_adapt = {
		('G6', 'I7'): ('G6', 'I7'),		# Dis1
		('G14', 'I16'): ('G14', 'I16'), # Dis2
		('G23', 'I25'): ('G23', 'I25'), # Dis3
		('G32', 'I35'): ('G32', 'I35'), # Dis4
		('G42', 'I45'): ('G42', 'I45'), # Dis5
	}
	ranges_to_update = cells_if_not_need_to_adapt if not need_to_adapt else cells_if_need_to_adapt
	for key, value in ranges_to_update.items():
		talents = previous_worksheet.get_values(start=key[0], end=key[1], returnas='matrix')
		current_worksheet.update_values(crange=value[0] + ':' + value[1], values=talents)

	print('Actualizada tab: ' + current_worksheet.title)
	print('')

def _hide_unused_talents_in_single_worksheet(destiny_sheet, worksheet_name, worksheet_identifier, answers_role_sheet, answers_role_row, answers_role_column, date_to_append, title_row_in_brief, talents_amount, mode):
	"""Hide unused talents"""
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
		# should_show_by_exception_5 = worksheet_identifier == 'SM' and title.lower() == 'iniciativa / autonomía' and 'iniciativa y autonomía' in answers_cell
		if should_show_by_match or should_show_by_exception_1 or should_show_by_exception_2 or should_show_by_exception_3 or should_show_by_exception_4:
			show_from = list(filter(lambda each: each.value.startswith(worksheet_identifier + str(i + 1)), cells_with_titles))[0].row - 1
			show_to = worksheet_destiny.rows if i + 1 == len(worksheet_talents) else list(filter(lambda each: each.value.startswith(worksheet_identifier + str(i + 2)), cells_with_titles))[0].row - 1
			worksheet_destiny.show_dimensions(show_from, show_to, dimension="ROWS")
			print('Mostrando talento: ' + title)

	print('')

def _build_evaluation_form_in_single_worksheet(destiny_sheet, auto_evaluation_sheet, manager_evaluation_sheet, exchange_evaluation_sheet, worksheet_name, worksheet_identifier, answers_role_sheet, answers_role_row, answers_role_column, date_to_append, talents_amount, mode):
	"""Copies evaluated talents from a sheet to destiny."""
	# In case mode is EXCHANGE_EVALUATION it uses auto_evaluation_sheet and manager_evaluation_sheet.
	# In case mode is FIRST_EVALUATION it uses exchange_evaluation_sheet.
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
		# should_copy_by_exception_5 = worksheet_identifier == 'SM' and title.lower() == 'iniciativa / autonomía' and 'iniciativa y autonomía' in answers_cell
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

def check_specific_talents(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append, mode):
	"""check certain talents"""
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