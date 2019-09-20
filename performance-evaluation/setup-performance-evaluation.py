# -*- coding: utf-8 -*-
from constants import RID_EVALUATION, AUTO_EVALUATION, MANAGER_EVALUATION, \
	EXCHANGE_EVALUATION, FIRST_EVALUATION, TEMPLATE_AUXILIAR_FILE_KEY, \
	TEMPLATE_TALENT_FILE_KEY, OPERATION_COPY_TABS, OPERATION_BUILD_EVALUATION_FORM, \
	OPERATION_HIDE_TALENTS, OPERATION_COPY_ANSWERS, OPERATION_COPY_FEEDBACK, \
	OPERATION_CREATE_SHEET, operations_by_mode_dictionary, folders_dictionary

from google_sheets import set_google_credentials, create_sheet, get_sheet_by_key, \
	get_feedback_sheet, get_answers_role_sheet, get_answers_role_row

from google_drive import get_key_by_path

from operations import copy_tabs, build_evaluation_form, hide_unused_talents, \
	copy_answers_role, copy_feedback, check_specific_talents

from inputs import get_date_to_append, get_mode, get_title, get_rid_instance_to_append

from validations import copy_should_be_omitted

from utils import wait_for_quota_renewal, on_finish

set_google_credentials()
template_auxiliar_sheet = get_sheet_by_key(TEMPLATE_AUXILIAR_FILE_KEY)
template_talent_sheet = get_sheet_by_key(TEMPLATE_TALENT_FILE_KEY)
date_to_append = get_date_to_append()
mode = get_mode()

if mode in operations_by_mode_dictionary[OPERATION_CREATE_SHEET]:
	title = get_title()

if mode in operations_by_mode_dictionary[OPERATION_COPY_FEEDBACK]:
	feedback_sheet = get_feedback_sheet()

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
	copy_feedback(destiny_sheet, feedback_sheet, date_to_append)

if mode != RID_EVALUATION:
	check_specific_talents(destiny_sheet, answers_role_sheet, answers_role_row, date_to_append, mode)

on_finish()