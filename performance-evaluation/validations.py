import readchar

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