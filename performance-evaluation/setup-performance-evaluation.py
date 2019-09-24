# -*- coding: utf-8 -*-
from constants import *

from google_operations import google_authentication, create_file, open_file, \
	get_feedback_sheet, find_key, find_and_open_file

from operations import copy_tabs, build_evaluation_form, hide_unused_talents, \
	copy_answers_role, copy_feedback, check_specific_talents

from inputs import get_date_to_append, get_mode, get_title, get_rid_instance_to_append, get_answers_role_row

from validations import copy_should_be_omitted

from utils import wait_for_quota_renewal, on_finish

google_authentication()
template_auxiliar_sheet = open_file(TEMPLATE_AUXILIAR_FILE_KEY)
template_talent_sheet = open_file(TEMPLATE_TALENT_FILE_KEY)
mode = get_mode()
date_to_append = get_date_to_append()

if mode in operations_by_mode_dictionary[OPERATION_CREATE_SHEET]:
	title = get_title()

if mode in operations_by_mode_dictionary[OPERATION_COPY_FEEDBACK]:
	feedback_sheet = get_feedback_sheet()

mode_build_evaluation_form = mode in operations_by_mode_dictionary[OPERATION_BUILD_EVALUATION_FORM]
mode_hide_talents = mode in operations_by_mode_dictionary[OPERATION_HIDE_TALENTS]
mode_copy_answers = mode in operations_by_mode_dictionary[OPERATION_COPY_ANSWERS]

if mode_build_evaluation_form or mode_hide_talents or mode_copy_answers:
	answers_role_sheet = open_file(ANSWERS_ROLE_FILE_KEY)
	answers_role_row = get_answers_role_row(answers_role_sheet)

if file_actions[mode]['create']:
    # El modo me dice que tengo que crear un file
	if file_actions[mode]['create']['key']:
		key = file_actions[mode]['create']['key']
	else:
		input('Al hacer click en Enter, te listará las carpetas de desempeño para que selecciones:')
		key = find_key(file_actions[mode]['create']['parent_key'], True)
		if key == FOLDER_BOOLEAN:
			folder_title = get_title()
			key = create_file(FOLDER_TYPE, folder_title, file_actions[mode]['create']['parent_key'])
	destiny_sheet = create_file(SHEET_TYPE, title, key)
else:
	input('Al hacer click en Enter, te listará las carpetas de desempeño para que selecciones. Luego, debes seleccionar el archivo del informe:')
	parent_key = find_key(file_actions[mode]['search']['key'], False)
	key = find_key(parent_key, False)
	destiny_sheet = open_file(key)

rid_instance_to_append = get_rid_instance_to_append() if mode == RID_EVALUATION else None

if mode in operations_by_mode_dictionary[OPERATION_COPY_TABS]:
	if not copy_should_be_omitted(destiny_sheet, date_to_append):
		copy_tabs(destiny_sheet, template_auxiliar_sheet, template_talent_sheet, date_to_append, rid_instance_to_append, mode)
		if not mode == RID_EVALUATION:
			wait_for_quota_renewal()

if mode in operations_by_mode_dictionary[OPERATION_BUILD_EVALUATION_FORM]:
	auto_evaluation_sheet = find_and_open_file(AUTO_EVALUATION_FOLDER_KEY, AUTO_EVALUATION_NAME) if mode == EXCHANGE_EVALUATION else None
	manager_evaluation_sheet = find_and_open_file(MANAGER_EVALUATION_FOLDER_KEY, MANAGER_EVALUATION_NAME) if mode == EXCHANGE_EVALUATION else None
	exchange_evaluation_sheet = find_and_open_file(EXCHANGE_EVALUATION_FOLDER_KEY, EXCHANGE_EVALUATION_NAME) if mode == FIRST_EVALUATION else None
	build_evaluation_form(destiny_sheet, date_to_append, mode, auto_evaluation_sheet, manager_evaluation_sheet, exchange_evaluation_sheet, answers_role_sheet, answers_role_row)
	wait_for_quota_renewal()

if mode in operations_by_mode_dictionary[OPERATION_HIDE_TALENTS]:
	hide_unused_talents(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append, mode)

if mode in operations_by_mode_dictionary[OPERATION_COPY_ANSWERS]:
	copy_answers_role(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append, mode)

if mode in operations_by_mode_dictionary[OPERATION_COPY_FEEDBACK]:
	copy_feedback(destiny_sheet, feedback_sheet, date_to_append)

if mode != RID_EVALUATION:
	check_specific_talents(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append, mode)

on_finish()