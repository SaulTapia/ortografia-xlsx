import openpyxl
from .regex_dict_xlsx import expressions

alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
def fix_xlsx(filename, starting_cell, ending_cell, sheetname):
    '''Fixes common spanish grammar errors on xlsx files'''
    print('Empieza el script')
    starting_separator = 0
    ending_separator = 0
    for i, _ in enumerate(starting_cell):
        if _.isdigit():
            starting_separator = i
            print(starting_separator)
            break

    for i, _ in enumerate(ending_cell):
        if _.isdigit():
            ending_separator = i
            break

    starting_number_string = starting_cell[starting_separator:]
    ending_number_string = ending_cell[ending_separator:]
    if starting_number_string[0] == '0' or ending_number_string[0] == '0':
        return 'El número de celda no puede empezar en 0'
    

    try:
        starting_number = int(starting_number_string)
        ending_number = int(ending_number_string)
    except Exception as e:
        return 'No juntes números con letras!'
        
    if starting_number > ending_number:
        return 'El comienzo no puede estar después del final!'

    if ending_number - starting_number > 20000:
        return 'No haré tanto'

    starting_letters = starting_cell[:starting_separator]
    ending_letters = ending_cell[:ending_separator]

    if len(starting_letters) > 3 or len(ending_letters) > 3:
        return 'No voy a hacer tanto'

    if len(starting_letters) > len(ending_letters):
        return 'El comienzo no puede estar después del final!'

    if len(starting_letters) == len(ending_letters):
        for pair in zip(starting_letters, ending_letters):
            if pair[1] > pair[0]:
                break
            elif pair[0] > pair[1]:
                return 'El comienzo no puede estar después del final!'
    



    try:
        wb = openpyxl.load_workbook(filename)
        sheet = wb[sheetname]
    except expression as identifier:
        return 'Esa hoja no existe!'
    try:
        while True:
            print(f'Arreglando la columna {starting_letters}')

            for i in range(starting_number, ending_number + 1):
                value = sheet[f'{starting_letters}{i}'].value
                if type(value) is str:
                    value = value.lower()
                    for expression, word in expressions.items():
                        while expression.match(value):
                            value = expression.sub(fr'\g<1>{word}\g<3>', value)
                            

                    first_letter = value[0]
                    value = value[1:]
                    sheet[f'{starting_letters}{i}'].value = first_letter.upper() + value

            if starting_letters == ending_letters:
                break

            letter_index = 1
            while True:
                last_letter = starting_letters[letter_index * -1]
                if last_letter not in alphabet:
                    return 'Todas las letras deben ser mayúsculas, sin acentos y sin ñ'

                if last_letter == 'Z':
                    if letter_index == len(starting_letters):
                        starting_letters = 'AA' + starting_letters[1:]
                        break
                    if letter_index == 1:
                        starting_letters = starting_letters[:-1] + 'A'
                    else:
                        starting_letters = starting_letters[:letter_index * - 1] + 'A' + starting_letters[(letter_index * -1) + 1]

                else:
                    last_letter_index = alphabet.index(last_letter)
                    if letter_index == 1:
                        starting_letters= starting_letters[:letter_index * -1] + alphabet[last_letter_index + 1]
                    else:
                        starting_letters = starting_letters[:letter_index * -1] + alphabet[last_letter_index + 1] + starting_letters[(letter_index * -1) + 1:]

                    break

                letter_index += 1

    except Exception as e:
        return f'''Revisa bien como ingresaste las celdas\n
                Error: {e}'''
        
    wb.save(filename)
    return 0