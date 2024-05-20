import xlsxwriter
from os import startfile, system
from os import path
from openpyxl import utils

class Tex_to_Excel:
    def __init__(self, tex_file, cleanup=False):
        self.tex_file = path.abspath(tex_file)
        self.excel_file = path.abspath(tex_file.replace('.tex', '.xlsx'))
        self.excel_wb = xlsxwriter.Workbook(self.excel_file)
        self.excel_sheet = self.excel_wb.add_worksheet()

        form_def = self.excel_wb.add_format(
            {"italic": 0, "valign": "top", "font_name": "ISOCPEUR", "font_size": 12, "text_wrap": 1,
             "border": 1})
        form_sub = self.excel_wb.add_format(
            {"italic": 0, "font_script": 2, "font_name": "ISOCPEUR", "font_size": 16, "text_wrap": 1,
             "border": 1})
        form_super = self.excel_wb.add_format(
            {"italic": 0, "font_script": 1, "font_name": "ISOCPEUR", "font_size": 16, "text_wrap": 1,
             "border": 1})
        form_bold = self.excel_wb.add_format(
            {"italic": 0, "valign": "top", "bold": 1, "font_name": "ISOCPEUR", "font_size": 12, "text_wrap": 1,
             "border": 1})
        form_sub_bold = self.excel_wb.add_format(
            {"italic": 0, "font_script": 2, "bold": 1, "font_name": "ISOCPEUR", "font_size": 16, "text_wrap": 1,
             "border": 1})
        form_super_bold = self.excel_wb.add_format(
            {"italic": 0, "font_script": 1, "bold": 1, "font_name": "ISOCPEUR", "font_size": 16, "text_wrap": 1,
             "border": 1})
        self.form_bundle = [form_def, form_sub, form_super, form_bold, form_sub_bold, form_super_bold]

        self.cleanup = cleanup

    def excel_address(self, row, column):
        return f"{utils.cell.get_column_letter(column)}{row}"

    def parce_Tex(self):
        with open(self.tex_file, encoding = 'UTF-8') as file:
            text = []
            for line in file:
                text.append(line.strip())
        file.close()
        return text

    def write_Excel(self):

        parced_file = self.parce_Tex()

        self.excel_sheet.set_column(0, 0, 200)

        for row, line in enumerate(parced_file):
            if self.cleanup == True:
                line = self.cleanup_crew(line)
            rich_string = self.get_rich_string(line)
            if rich_string == None:
                self.excel_sheet.write(row, 0, line, self.form_bundle[0])
            elif type(rich_string) == xlsxwriter.format.Format:
                self.excel_sheet.write(row, 0, line[1:-1], rich_string)
            else:
                self.excel_sheet.write_rich_string(row, 0, *rich_string)

        try:
            self.excel_wb.close()
        except xlsxwriter.exceptions.FileCreateError:
            return

        startfile(self.excel_file)
        #system("start " + "\"\" \"" + self.excel_file + "\"")

    def cleanup_crew(self, string):
        # Find the result part and replace "·" with " " so it looks better
        end_pos = string.rfind('=') # Find the last "="
        if end_pos != -1: # If it exists
            if string.find('·', end_pos) != -1: # Find the first "·" in the result
                string = string[:end_pos] + string[end_pos:].replace('·', ' ', 1) # And replace it once
        return string

        """
        def split_pos(string):  # Get the positions of "=" sings
            pos1 = string.find('=')  # Find the first =
            if pos1 == -1:
                return None  # If no = return None
            start = pos1 + 1
            pos2 = string.find('=', start)  # Find the second =
            if pos2 == -1:
                return None  # If no second = return None
            start = pos2 + 1
            pos3 = string.find('=', start)  # Find the third =
            if pos3 == -1:
                return None  # If no third = return None
            return pos1, pos2, pos3

        def skip_units(substit):
            string_no_units = ''  # Epty wtring for writing in
            operators = ('/', '·', '+', '-', ' ', '(', ')')  # The list of operators which mean the unit has ended
            found_unit = False  # The flag that the unit was found
            for i in range(len(substit) - 1):
                char = substit[i]
                next_char = substit[i + 1]
                if char in operators and next_char.isalpha():  # If there's an operation with a letter that counts as a unit
                    found_unit = True
                if found_unit == True:
                    if next_char not in operators:
                        continue  # If flag is set and there're no operators skip the char
                    else:
                        found_unit = False  # If there's an operator next then reset the flag
                else:
                    string_no_units += char
            string_no_units += ' '  # Add the space in the end

            return string_no_units

        def restore_parens(defin, substit):
            operators = ('/', '·', '+', '-', '(', ')')  # What is consodered to be in operator

            defin = '(' + defin.strip(' = ').replace(' ', '') + ')'
            substit = '(' + substit.strip(' = ').replace(' ', '') + ')'
            # Add () to make it easier to find the first (. Get rid of = and spaces

            ops_defin = []  # A list of operations and other stuff in defin
            stuff_found = True  # Set as True because we added () at the start
            for char in defin:
                if char not in operators:  # If we found some stuff that isn't an operator
                    if stuff_found == False:  # And we didn't add '' already
                        stuff_found = True  # Set the flag
                        ops_defin.append('')  # Add ''
                if char in operators:  # If we found an operator
                    if stuff_found == True:  # And the last elem wasn't an operator
                        ops_defin.append(char)  # Add the operator as an element
                        stuff_found = False  # Reset the flag
                    else:
                        ops_defin[-1] = ops_defin[-1] + char
                        # If the last elem was an operator add it to the previous operator

            split_substit = []  # A list of text and operators in substit
            stuff_found = True  # Set as True because we added () at the start
            for char in substit:
                if char not in operators:  # If we found text
                    if stuff_found == False:  # And we didn't add the new text elem
                        split_substit.append(char)  # Add the new text elem
                        stuff_found = True  # Set the flag
                    else:
                        split_substit[-1] += char
                        # If the last char was text also just add to it
                if char in operators:  # If we found an operator
                    if stuff_found == True:  # And the last elem wasn't an operator
                        split_substit.append(char)  # Add the new operator elem
                        stuff_found = False  # Set the flag
                    else:
                        split_substit[-1] = split_substit[-1] + char
                        # If the last elem was an operator just add to it

            for i in range(len(ops_defin)):
                if ops_defin[i] != '':
                    split_substit[i] = ops_defin[i]
            # For every operator in ops_defin replace the corresponding operator in split_subsist

            string_with_parens = ''.join(split_substit)  # Make a string
            string_with_parens = string_with_parens[1:-1]  # Cut the () we added at the beginning
            string_with_parens = string_with_parens.replace('+', ' + ')  # Add back the spaces
            string_with_parens = string_with_parens.replace('-', ' - ')
            string_with_parens = "= " + string_with_parens + ' '  # Add back the =

            return string_with_parens

        def cleanup_parens(substit):
            digit = False  # Flag for a digit
            parens = False  # Flag for a parens
            replace = []  # Empty list for digits that are surrounded by ()
            lone_digit = ''  # An empty digit
            for char in substit:
                if char == '(':  # Set the flag that the parens are opened
                    parens = True
                elif char == ')':  # Parens are closed - reset the flag
                    parens = False
                    if lone_digit:  # If there was a digit between ()
                        replace.append(lone_digit)  # Add it to the replace list
                    lone_digit = ''  # Reset the digit
                elif char.isdigit() or char == ',':  # If char is digit or a comma it's a part of a number
                    digit = True  # Set the flag
                else:
                    digit = False  # If not where's an operator and we can't strip parens
                    parens = False  # So reset both flags
                    lone_digit = ''  # And reset the digit
                if digit and parens:  # If both flags are present
                    lone_digit += char  # Add the digit to the pile

            for number in replace:
                substit = substit.replace("(" + number + ")", number)
                # Replace all numbers (number) with number

            return substit

        positions = split_pos(string)
        if positions:  # Check if cleanup is needed
            pos1, pos2, pos3 = positions  # Unpack the positions
            defin = string[pos1:pos2]  # The definition part of the string
            substit = string[pos2:pos3]  # The intermediary calculation part or substitution
            string_no_units = skip_units(substit)  # First remove units because other stuff doesn't work with them
            string_with_parens = restore_parens(defin, string_no_units)  # Then add back the ()
            clean_string = cleanup_parens(string_with_parens)  # And finally clean up extra ()
        else:
            return string  # If not return the given string

        return string[:pos2] + clean_string + string[pos3:]  # Combine the clean substitution with no units with the others
        """

    def get_rich_string(self, line):
        """Возвращает готовые переменные в виде 'строка 1', 'стиль 1' ... 'строка N', 'стиль N' для размеченной строки"""
        symbols = ['|', '?', '^']
        form_def, form_sub, form_super, form_bold, form_sub_bold, form_super_bold = self.form_bundle
        equation = line

        def style(s1, s2):  # Комбинирование стилей
            if s2 not in s1:
                return s1 + s2
            else:
                return s1.replace(s2, "")

        def style_to_form(symbol):  # Перевод стиля в форму
            try:
                styles = {"": form_def, "|": form_sub, "^": form_super,
                      "?": form_bold, "?|": form_sub_bold, "?^": form_super_bold,
                      "|?": form_sub_bold, "^?": form_super_bold}
                form = styles[symbol]
            except KeyError:
                return
            return form

        def len_interval(interval):  # Длина интервала
            return interval[1] - interval[0]

        if all(equation.find(symbol) == -1 for symbol in symbols):  # Проверка, есть ли символы разметки
            return

        # Формирование строки без символов и позиций символов
        string = ''
        i = 0
        k = 0
        positions = []

        while any(equation.find(symbol) != -1 for symbol in symbols):
            if equation[i] in symbols:
                k += i
                positions.append([k, equation[i]])
                string += equation[:i]
                equation = equation[i + 1:]
                i = 0
            else:
                i += 1
        string += equation

        # Если символов не четное количество
        if len(positions) % 2 != 0:
            return

        # Формирование интервалов
        intervals = [[0, positions[0][0], ""]]
        for i in range(len(positions) - 1):
            prev = intervals[i]
            next_pos = positions[i + 1]
            interval = [positions[i][0], next_pos[0], style(prev[2], positions[i][1])]
            intervals.append(interval)
        if intervals[-1][1] != len(string):
            intervals.append([intervals[-1][1], len(string), ''])
        empty_intervals = [interval for interval in intervals if len_interval(interval) == 0]
        for empty in empty_intervals:
            intervals.remove(empty)

        # Перевод стилей в формы
        for interval in intervals:
            interval[2] = style_to_form(interval[2])
            if interval[2] == None:
                return line

        # Если интервал всего один, то возвращается стиль
        if len(intervals) == 1:
            return intervals[0][2]

        replace_symbols = {'__': '+', '..': ' ', 'ЗI': '3I', '_': '-'}  # A dict of subscript symbols for replacement

        rich_string = []
        for interval in intervals:
            rich_string.append(interval[2])
            temp_string = string[interval[0]:interval[1]] # Create a temp string for convenience
            for symbol in replace_symbols.keys():
                # Replace all the symbols
                temp_string = temp_string.replace(symbol, replace_symbols[symbol])
            if any([interval[2] == form_sub, interval[2] == form_sub_bold]):  # If handling a subscript
                rich_string.append(temp_string.upper()) # Make it all uppercase
            else:
                rich_string.append(temp_string) # If not, just add the string
        rich_string.append(form_def)

        return rich_string
