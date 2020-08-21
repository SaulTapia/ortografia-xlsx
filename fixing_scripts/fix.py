import openpyxl
from .regex_dict_xlsx import expressions

alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
def fix_xlsx(filename, starting_cell, ending_cell, sheetname):
    '''Fixes common spanish grammar errors on xlsx files'''
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

    starting_number = int(starting_cell[starting_separator:])
    ending_number = int(ending_cell[ending_separator:])
        
    if starting_number > ending_number:
        return 1

    starting_letters = starting_cell[:starting_separator]
    ending_letters = ending_cell[:ending_separator]
    if len(starting_letters) > len(ending_letters):
        return 1
    elif len(starting_letters) == len(ending_letters):
        aux_starting = list(starting_letters)
        aux_ending = list(ending_letters)
        aux_starting.reverse()
        aux_ending.reverse()
        for pair in zip(aux_starting, aux_ending):
            if pair[0] > pair[1]:
                return 1

    wb = openpyxl.load_workbook(filename)
    sheet = wb[sheetname]
        

    #sheet['A1'].value = 'Hola'
    # wb.save(filename)

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
            if last_letter == 'Z':
                if letter_index == len(starting_letters):
                    starting_letters = 'AA' + starting_letters[1:]
                    break
                if letter_index == 1:
                    starting_letters = starting_letters[letter_index * - 1] + 'A'
                else:
                    starting_letters = starting_letters[letter_index * - 1] + 'A' + starting_letters[(letter_index * -1) + 1]

            else:
                last_letter_index = alphabet.index(last_letter)
                if letter_index == 1:
                    starting_letters= starting_letters[:letter_index * -1] + alphabet[last_letter_index + 1]
                else:
                    starting_letters = starting_letters[:letter_index * -1] + alphabet[last_letter_index + 1] + starting_letters[(letter_index * -1) + 1:]

                break

            letter_index += 1
        
    wb.save(filename)
    return 0