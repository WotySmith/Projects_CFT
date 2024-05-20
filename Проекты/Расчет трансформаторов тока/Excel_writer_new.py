import xlsxwriter
from os import startfile
from tkinter import messagebox

def str_cm(num_float):
    str_dot = str(round(num_float,2))
    str_cm = str_dot.replace(".", ",")
    if str_cm.endswith(",0") == True:
        str_cm = str_cm.replace(",0", "")
    return str_cm

def get_rich_string(equation, form_sub, form_super, form_def, form_bold):
    rich_string = []
    i = 0
    while len(equation) > 0:
        if equation.find("|") == -1 and equation.find("^") == -1 and equation.find("?") == -1:
            rich_string.append(form_def)
            rich_string.append(equation)
            rich_string.append(form_def)
            return rich_string
                    
        if equation[i] == "|" or equation[i] == "^" or equation[i] == "?":
            pos_start = i
            symbol = equation[i]
            pos_stop = equation.find(symbol, pos_start+1)
            alt_str = equation[pos_start+1:pos_stop]
            equation = equation[:pos_start] + equation[pos_stop+1:]
            if pos_start > 0:
                def_str = equation[:pos_start]
                equation = equation[pos_start:]
                rich_string.append(form_def)
                rich_string.append(def_str)
            if symbol == "|":
                rich_string.append(form_sub)
            elif symbol == "^":
                rich_string.append(form_super)
            elif symbol == "?":
                rich_string.append(form_bold)
            rich_string.append(alt_str)
            i = 0
        else:
            i += 1
    rich_string.append(form_def)
    return rich_string

def save_report(data, italic):

    Num_TT_mult = data["Num_TT_mult"]

    workbook = xlsxwriter.Workbook("log_fixed.xlsx")
    worksheet = workbook.add_worksheet()

    form_super = workbook.add_format({"italic": italic, "font_script": 1, "font_name": "ISOCPEUR", "font_size": 16})
    form_super_bold = workbook.add_format({"italic": italic, "bold":1, "font_script": 1, "font_name": "ISOCPEUR", "font_size": 16})
    form_sub = workbook.add_format({"italic": italic, "font_script": 2, "font_name": "ISOCPEUR", "font_size": 16})
    form_def = workbook.add_format({"italic": italic, "valign": "vcenter", "font_name": "ISOCPEUR", "font_size": 12, "text_wrap":1, "border": 1})
    form_def_rotated = workbook.add_format({"italic": italic, "valign": "vcenter", "rotation": 90, "font_name": "ISOCPEUR", "font_size": 12, "text_wrap":1, "border": 1})
    form_bold = workbook.add_format({"italic": italic, "valign": "vcenter", "bold": 1, "font_name": "ISOCPEUR", "font_size": 12, "border": 1})
    form_bold_center = workbook.add_format({"italic": italic, "bold": 1, "align": "center", "valign": "vcenter", "font_name": "ISOCPEUR", "font_size": 12, "border": 1})
    form_bundle = [form_sub, form_super, form_def, form_bold]

    worksheet.set_column(0, 0, 35)
    worksheet.set_column(1, 1, 40)
    worksheet.set_column(2, 4, 25)

    row = 1

    #Параметры ТТ
    worksheet.write(f"A{row}", "Номинальный первичный ток", form_def)
    worksheet.write_rich_string(f"B{row}", *get_rich_string(
        "I|1|, А", *form_bundle))
    worksheet.write(f"C{row}", data["I1"], form_def)
    row += 1
    
    worksheet.write(f"A{row}", "Номинальный вторичный ток", form_def)
    worksheet.write_rich_string(f"B{row}", *get_rich_string(
        "I|2|, А", *form_bundle))
    worksheet.write(f"C{row}", data["I2"], form_def)
    row += 1
    
    worksheet.write(f"A{row}", "Класс точности", form_def)
    worksheet.write(f"B{row}", "Кл.точн.", form_def)
    if data["k_gamma"] > 0.1:
        worksheet.write(f"C{row}", "10P", form_def)
    else:
        worksheet.write(f"C{row}", "10PR", form_def)
    row += 1
    
    worksheet.write(f"A{row}", "Номинальная нагрузка", form_def)
    worksheet.write_rich_string(f"B{row}", *get_rich_string(
        "S|НОМ|, ВА", *form_bundle))
    worksheet.write(f"C{row}", str_cm(data["S_ras"]), form_def)
    row += 1

    worksheet.write(f"A{row}", "Номинальная предельная кратность", form_def)
    worksheet.write_rich_string(f"B{row}", *get_rich_string(
        "K|НОМ|", *form_bundle))
    worksheet.write(f"C{row}", str_cm(data["K_ras"]), form_def)
    row += 1
    
    if data["X2tt"] == 0:
        worksheet.write(f"A{row}", "Активное сопротивление вторичной обмотки пост. току", form_def)
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            "R|2ТТ|, Ом", *form_bundle))
    else:
        worksheet.write(f"A{row}", "Cопротивление вторичной обмотки", form_def)
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            "R|2ТТ|, Ом\n"+
            "X|2ТТ|, Ом", *form_bundle))
    if data["X2tt"] == 0:
        worksheet.write(f"C{row}", data["R2tt"], form_def)
    else:
        worksheet.write(f"C{row}", str_cm(data["R2tt"]) + "\n" + str_cm(data["X2tt"]), form_def)    
    row += 1

    #Параметры 3-ф КЗ
    if data["ignore_Int3"] == False:
        worksheet.write(f"A{row}", f"Максимальный ток 3-ф КЗ по ТТ при КЗ в зоне действия защиты", form_def)
        worksheet.write(f"B{row}", f"Внутреннее 3-ф КЗ", form_def)
        worksheet.write_rich_string(f"C{row}", *get_rich_string(f"I|ВНУТР|^(3)^ = {str_cm(data['Ikz3_Int'])} А", *form_bundle))
        row += 1

    if data["ignore_Ext3"] == False:
        worksheet.write(f"A{row}", f"Максимальный ток 3-ф КЗ по ТТ при КЗ вне зоны действия защиты", form_def)
        worksheet.write(f"B{row}", f"Внешнее 3-ф КЗ", form_def)    
        worksheet.write_rich_string(f"C{row}", *get_rich_string(f"I|ВНЕШ|^(3)^ = {str_cm(data['Ikz3_Ext'])} А", *form_bundle))
        row += 1

    if data["ignore_Int3"] == False:
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"T|p ВНУТР| - постоянная затухания апериод. сост. тока 3-ф КЗ в зоне действия защиты", *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string("T|p ВНУТР|^(3)^ = x|1| /(\u0277R|1|)", *form_bundle))
        worksheet.write_rich_string(f"C{row}", *get_rich_string(f"T|p ВНУТР|^(3)^ = {str_cm(data['Tp3_Int'])} мс", *form_bundle))
        row += 1

    if data["ignore_Ext3"] == False:
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"T|p ВНЕШ| - постоянная затухания апериод. сост. тока 3-ф КЗ вне зоны действия защиты", *form_bundle))  
        worksheet.write_rich_string(f"B{row}", *get_rich_string("T|p ВНЕШ|^(3)^ = x|1| /(\u0277R|1|)", *form_bundle))    
        worksheet.write_rich_string(f"C{row}", *get_rich_string(f"T|p ВНЕШ|^(3)^ = {str_cm(data['Tp3_Ext'])} мс", *form_bundle))
        row += 1

    #Параметры 1-ф КЗ
    if data["ignore_Int1"] == False:
        worksheet.write(f"A{row}", f"Максимальный ток 1-ф КЗ по ТТ при КЗ в зоне действия защиты", form_def)
        worksheet.write(f"B{row}", f"Внутреннее 1-ф КЗ", form_def)
        worksheet.write_rich_string(f"C{row}", *get_rich_string(f"I|ВНУТР|^(1)^ = {str_cm(data['Ikz1_Int'])} А", *form_bundle))
        row += 1

    if data["ignore_Ext1"] == False:
        worksheet.write(f"A{row}", f"Максимальный ток 1-ф КЗ по ТТ при КЗ вне зоны действия защиты", form_def)
        worksheet.write(f"B{row}", f"Внешнее 1-ф КЗ", form_def)    
        worksheet.write_rich_string(f"C{row}", *get_rich_string(f"I|ВНЕШ|^(1)^ = {str_cm(data['Ikz1_Ext'])} А", *form_bundle))
        row += 1

    if data["ignore_Int1"] == False:
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"T|p ВНУТР| - постоянная затухания апериод. сост. тока 1-ф КЗ в зоне действия защиты", *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string(f"T|p ВНУТР|^(1)^ = (2x|1|+x|0|) /(\u0277(2R|1|+R|0|))", *form_bundle))
        worksheet.write_rich_string(f"C{row}", *get_rich_string(f"T|p ВНУТР|^(1)^ = {str_cm(data['Tp1_Int'])} мс", *form_bundle))
        row += 1

    if data["ignore_Ext1"] == False:
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"T|p ВНЕШ| - постоянная затухания апериод. сост. тока 1-ф КЗ вне зоны действия защиты", *form_bundle))    
        worksheet.write_rich_string(f"B{row}", *get_rich_string(f"T|p ВНЕШ|^(1)^ = (2x|1|+x|0|) /(\u0277(2R|1|+R|0|))", *form_bundle))    
        worksheet.write_rich_string(f"C{row}", *get_rich_string(f"T|p ВНЕШ|^(1)^ = {str_cm(data['Tp1_Ext'])} мс", *form_bundle))
        row += 1
          
    #Расчетная нагрузка
    worksheet.merge_range(f"A{row}:C{row}", "Определение расчетной нагрузки", form_bold_center)
    row += 1
    
    worksheet.merge_range(f"A{row}:B{row}", "", form_def)
    worksheet.write_rich_string(f"A{row}", *get_rich_string("S - сечение контрольного кабеля, мм^2^", *form_bundle))
    worksheet.write_rich_string(f"C{row}", form_bold, "S = "+str_cm(data["Pop_ras"])+" мм", form_super_bold, "2", form_bold)
    row += 1
    
    worksheet.write_rich_string(f"A{row}", *get_rich_string("R|КАБ| - сопротивление жил контрольного кабеля, Ом", *form_bundle))
    worksheet.write_rich_string(f"B{row}", *get_rich_string(
        "R|КАБ| = \u03c1l/S;\n" +
        "\u03c1 = 1/57 (медь)\n" +
        f"l = {str_cm(data['L_cab'])} м - длина контрольного кабеля", *form_bundle))
    worksheet.write_rich_string(f"C{row}", *get_rich_string(
        f"R|КАБ| = {str_cm(data['R_cab'])} Ом", *form_bundle))
    row += 1

    # Активная нагрузка
    if data["X_rele"] == 0:
        if (data["ignore_Int3"] == False) or (data["ignore_Ext3"] == False):
            worksheet.write_rich_string(f"A{row}", *get_rich_string(
                f"R|НАГР.РАСЧ|^(3)^ - расчетное сопротивление нагрузки при 3-ф КЗ, Ом", *form_bundle))
            if data["CON_type"] == "star":
                worksheet.write_rich_string(f"B{row}", *get_rich_string(
                    "R|НАГР.РАСЧ|^(3)^ = " + Num_TT_mult[:-1] + "R|КАБ| + R|ТЕРМ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                    f"R|ТЕРМ| = S/I|2|^2^ = {str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^ = {str_cm(data['R_rele'])} Ом;\n" +
                    f"R|ПЕР| = {str_cm(data['Rp'])} Ом", *form_bundle))
            if data["CON_type"] == "triangle":
                worksheet.write_rich_string(f"B{row}", *get_rich_string(
                    "R|НАГР.РАСЧ|^(3)^ = " + Num_TT_mult[:-1] + "3R|КАБ| + 3R|ТЕРМ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                    f"R|ТЕРМ| = S/I|2|^2^ = {str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^ = {str_cm(data['R_rele'])} Ом;\n" +
                    f"R|ПЕР| = {str_cm(data['Rp'])} Ом", *form_bundle))
            if data["CON_type"] == "part Y":
                worksheet.write_rich_string(f"B{row}", *get_rich_string(
                    "R|НАГР.РАСЧ|^(3)^ = " + Num_TT_mult[:-1] + "\u221a3R|КАБ| + 2R|ТЕРМ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                    f"R|ТЕРМ| = S/I|2|^2^ = {str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^ = {str_cm(data['R_rele'])} Ом;\n" +
                    f"R|ПЕР| = {str_cm(data['Rp'])} Ом", *form_bundle))
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"R|НАГР.РАСЧ|^(3)^ = {str_cm(data['R3'])} Ом", *form_bundle))
            row += 1

        if (data["ignore_Int1"] == False) or (data["ignore_Ext1"] == False):
            worksheet.write_rich_string(f"A{row}", *get_rich_string(
                f"R|НАГР.РАСЧ|^(1)^ - расчетное сопротивление нагрузки при 1-ф КЗ, Ом", *form_bundle))        
            if data["CON_type"] == "star":
                worksheet.write_rich_string(f"B{row}", *get_rich_string(
                    "R|НАГР.РАСЧ|^(1)^ = " + Num_TT_mult[:-1] + "2R|КАБ| + R|ТЕРМ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                    f"R|ТЕРМ| = S/I|2|^2^ = {str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^ = {str_cm(data['R_rele'])} Ом;\n" +
                    f"R|ПЕР| = {str_cm(data['Rp'])} Ом", *form_bundle))
            if data["CON_type"] == "triangle":
                worksheet.write_rich_string(f"B{row}", *get_rich_string(
                    "R|НАГР.РАСЧ|^(1)^ = " + Num_TT_mult[:-1] + "2R|КАБ| + 2R|ТЕРМ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                    f"R|ТЕРМ| = S/I|2|^2^ = {str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^ = {str_cm(data['R_rele'])} Ом;\n" +
                    f"R|ПЕР| = {str_cm(data['Rp'])} Ом", *form_bundle))
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"R|НАГР.РАСЧ|^(1)^ = {str_cm(data['R1'])} Ом", *form_bundle))
            row += 1

        if data["calc10_2ph"] == True:
            worksheet.write_rich_string(f"A{row}", *get_rich_string(
                "R|НАГР.РАСЧ|^(2)^ - расчетное сопротивление нагрузки при 2-ф КЗ, Ом", *form_bundle))
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                "R|НАГР.РАСЧ|^(2)^ = " + Num_TT_mult[:-1] + "2R|КАБ| + 2R|ТЕРМ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                f"R|ТЕРМ| = S/I|2|^2^ = {str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^ = {str_cm(data['R_rele'])} Ом;\n" +
                f"R|ПЕР| = {str_cm(data['Rp'])} Ом", *form_bundle))
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"R|НАГР.РАСЧ|^(2)^ = {str_cm(data['R2'])} Ом", *form_bundle))
            row += 1

    # Активно-индуктивная нагрузка
    else:
        if (data["ignore_Int3"] == False) or (data["ignore_Ext3"] == False):
            worksheet.write_rich_string(f"A{row}", *get_rich_string(
                f"Z|НАГР.РАСЧ|^(3)^ - расчетное сопротивление нагрузки при 3-ф КЗ, Ом", *form_bundle))
            if data["CON_type"] == "star":
                worksheet.write_rich_string(f"B{row}", *get_rich_string(
                    "R|НАГР.РАСЧ|^(3)^ = " + Num_TT_mult[:-1] + "R|КАБ| + R|РЕЛЕ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                    f"R|РЕЛЕ| = S/I|2|^2^cos\u03c6|РЕЛЕ| = {str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^*{str_cm(data['cosf_rele'])}"+
                    f" = {str_cm(data['R_rele'])} Ом;\n" +
                    f"R|ПЕР| = {str_cm(data['Rp'])} Ом;\n"+
                    f"R|НАГР.РАСЧ|^(3)^ = {str_cm(data['R3'])} Ом;\n" +
                    "X|НАГР.РАСЧ| = " + Num_TT_mult[:-2] + "X|РЕЛЕ| = " + Num_TT_mult[:-2] + "S/I|2|^2^sin\u03c6|РЕЛЕ| = "+
                    Num_TT_mult[:-2] + f"{str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^*{str_cm(data['sinf_rele'])}"+
                    f" = {str_cm(data['X_rele'])} Ом;\n" +
                    "Z|НАГР.РАСЧ|^(3)^ = \u221a(R|НАГР.РАСЧ|^2^ + X|НАГР.РАСЧ|^2^)", *form_bundle))
            if data["CON_type"] == "triangle":
                worksheet.write_rich_string(f"B{row}", *get_rich_string(
                    "R|НАГР.РАСЧ|^(3)^ = " + Num_TT_mult[:-1] + "3R|КАБ| + 3R|РЕЛЕ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                    f"R|РЕЛЕ| = S/I|2|^2^cos\u03c6|РЕЛЕ| = {str_cm(data['S_rele'])}/{data(data['I2'])}^2^*{str_cm(data['cosf_rele'])}"+
                    f" = {str_cm(data['R_rele'])} Ом;\n" +
                    "R|ПЕР| = {str_cm(data['Rp'])} Ом;\n"+
                    f"R|НАГР.РАСЧ|^(3)^ = {str_cm(data['R3'])} Ом;\n" +
                    "X|НАГР.РАСЧ| = " + Num_TT_mult[:-2] + "3X|РЕЛЕ| = " + Num_TT_mult[:-2] + "3*S/I|2|^2^sin\u03c6|РЕЛЕ| = "+
                    Num_TT_mult[:-2] + f"3*{str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^*{str_cm(data['sinf_rele'])}"+
                    f" = {str_cm(data['X_rele'])} Ом;\n" +
                    "Z|НАГР.РАСЧ|^(3)^ = \u221a(R|НАГР.РАСЧ|^2^ + X|НАГР.РАСЧ|^2^)", *form_bundle))
            if data["CON_type"] == "part 2":
                worksheet.write_rich_string(f"B{row}", *get_rich_string(
                    "R|НАГР.РАСЧ|^(3)^ = " + Num_TT_mult[:-1] + "\u221a3R|КАБ| + 2R|РЕЛЕ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                    f"R|РЕЛЕ| = S/I|2|^2^cos\u03c6|РЕЛЕ| = {str_cm(data['S_rele'])}/{data(data['I2'])}^2^*{str_cm(data['cosf_rele'])}"+
                    f" = {str_cm(data['R_rele'])} Ом;\n" +
                    "R|ПЕР| = {str_cm(data['Rp'])} Ом;\n"+
                    f"R|НАГР.РАСЧ|^(3)^ = {str_cm(data['R3'])} Ом;\n" +
                    "X|НАГР.РАСЧ| = " + Num_TT_mult[:-2] + "2X|РЕЛЕ| = " + Num_TT_mult[:-2] + "2*S/I|2|^2^sin\u03c6|РЕЛЕ| = "+
                    Num_TT_mult[:-2] + f"2*{str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^*{str_cm(data['sinf_rele'])}"+
                    f" = {str_cm(data['X_rele'])} Ом;\n" +
                    "Z|НАГР.РАСЧ|^(3)^ = \u221a(R|НАГР.РАСЧ|^2^ + X|НАГР.РАСЧ|^2^)", *form_bundle))            
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"Z|НАГР.РАСЧ|^(3)^ = {str_cm(data['Z_nagr3'])} Ом", *form_bundle))
            row += 1

        if (data["ignore_Int1"] == False) or (data["ignore_Ext1"] == False):
            worksheet.write_rich_string(f"A{row}", *get_rich_string(
                f"Z|НАГР.РАСЧ|^(1)^ - расчетное сопротивление нагрузки при 1-ф КЗ, Ом", *form_bundle))
            if data["CON_type"] == "star":
                worksheet.write_rich_string(f"B{row}", *get_rich_string(
                    "R|НАГР.РАСЧ|^(1)^ = " + Num_TT_mult[:-1] + "2R|КАБ| + R|РЕЛЕ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                    f"R|РЕЛЕ| = S/I|2|^2^cos\u03c6|РЕЛЕ| = {str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^*{str_cm(data['cosf_rele'])}"+
                    f" = {str_cm(data['R_rele'])} Ом;\n" +
                    "R|ПЕР| = {str_cm(data['Rp'])} Ом;\n"+
                    f"R|НАГР.РАСЧ|^(1)^ = {str_cm(data['R1'])} Ом;\n" +
                    "X|НАГР.РАСЧ| = " + Num_TT_mult[:-2] + "X|РЕЛЕ| = " + Num_TT_mult[:-2] + "S/I|2|^2^sin\u03c6|РЕЛЕ| = "+
                    Num_TT_mult[:-2] + f"{str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^*{str_cm(data['sinf_rele'])}"+
                    f" = {str_cm(data['X_rele'])} Ом;\n" +
                    "Z|НАГР.РАСЧ|^(1)^ = \u221a(R|НАГР.РАСЧ|^2^ + X|НАГР.РАСЧ|^2^)", *form_bundle))
            if data["CON_type"] == "triangle":
                worksheet.write_rich_string(f"B{row}", *get_rich_string(
                    "R|НАГР.РАСЧ|^(1)^ = " + Num_TT_mult[:-1] + "2R|КАБ| + 2R|РЕЛЕ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                    f"R|РЕЛЕ| = S/I|2|^2^cos\u03c6|РЕЛЕ| = {str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^*{str_cm(data['cosf_rele'])}"+
                    f" = {str_cm(data['R_rele'])} Ом;\n" +
                    "R|ПЕР| = {str_cm(data['Rp'])} Ом;\n"+
                    f"R|НАГР.РАСЧ|^(1)^ = {str_cm(data['R1'])} Ом;\n" +
                    "X|НАГР.РАСЧ| = " + Num_TT_mult[:-2] + "2X|РЕЛЕ| = " + Num_TT_mult[:-2] + "2*S/I|2|^2^sin\u03c6|РЕЛЕ| = "+
                    Num_TT_mult[:-2] + f"2*{str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^*{str_cm(data['sinf_rele'])}"+
                    f" = {str_cm(data['X_rele'])} Ом;\n" +
                    "Z|НАГР.РАСЧ|^(1)^ = \u221a(R|НАГР.РАСЧ|^2^ + X|НАГР.РАСЧ|^2^)", *form_bundle))
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"Z|НАГР.РАСЧ|^(1)^ = {str_cm(data['Z_nagr1'])} Ом", *form_bundle))
            row += 1

        if data["calc10_2ph"] == True:
            worksheet.write_rich_string(f"A{row}", *get_rich_string(
                "Z|НАГР.РАСЧ|^(2)^ - расчетное сопротивление нагрузки при 2-ф КЗ, Ом", *form_bundle))
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                "R|НАГР.РАСЧ|^(2)^ = " + Num_TT_mult[:-1] + "2R|КАБ| + 2R|РЕЛЕ| + R|ПЕР|" + Num_TT_mult[-1:] + ";\n" +
                f"R|РЕЛЕ| = S/I|2|^2^cos\u03c6|РЕЛЕ| = {str_cm(data['S_rele'])}/{data(data['I2'])}^2^*{str_cm(data['cosf_rele'])}"+
                f" = {str_cm(data['R_rele'])} Ом;\n" +
                "R|ПЕР| = {str_cm(data['Rp'])} Ом;\n"+
                f"R|НАГР.РАСЧ|^(2)^ = {str_cm(data['R2'])} Ом;\n" +
                "X|НАГР.РАСЧ| = " + Num_TT_mult[:-2] + "2X|РЕЛЕ| = " + Num_TT_mult[:-2] + "2*S/I|2|^2^sin\u03c6|РЕЛЕ| = "+
                Num_TT_mult[:-2] + f"2*{str_cm(data['S_rele'])}/{str_cm(data['I2'])}^2^*{str_cm(data['sinf_rele'])}"+
                f" = {str_cm(data['X_rele'])} Ом;\n" +
                "Z|НАГР.РАСЧ|^(2)^ = \u221a(R|НАГР.РАСЧ|^2^ + X|НАГР.РАСЧ|^2^)", *form_bundle))            
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"Z|НАГР.РАСЧ|^(2)^ = {str_cm(data['Z_nagr2'])} Ом", *form_bundle))
            row += 1

    #Номинальная нагрузка
    worksheet.merge_range(f"A{row}:C{row}", "Определение номинальной нагрузки", form_bold_center)
    row += 1
    
    worksheet.write_rich_string(f"A{row}", *get_rich_string(
        "Z|НАГР.НОМ|, Ом", *form_bundle))
    worksheet.write_rich_string(f"B{row}", *get_rich_string(
        "Z|НАГР.НОМ| = S|НОМ|/I|2|^2^", *form_bundle))
    worksheet.write_rich_string(f"C{row}", *get_rich_string(
        f"Z|НАГР.НОМ| = {str_cm(data['Z_ras'])} Ом", *form_bundle))
    row += 1

    #Проверка напряжения
    worksheet.merge_range(f"A{row}:C{row}", "Проверка допустимого напряжения на выводах вторичной обмотки ТТ", form_bold_center)
    row += 1

    #U2 при 3ф КЗ
    if (data["ignore_Int3"] == False) or (data["ignore_Ext3"] == False):
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"U|2MAX|^(3)^ - максимальное значение напряжения на выводах вторичной обмотки в режиме 3-ф КЗ, В",
            *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            f"U|2MAX|^(3)^ = \u221a2K|у|I|КЗ|Z|НАГР|/K|ТТ|;\n"
            "K|у| = 2;\n"
            f"U|2MAX|^(3)^ = \u221a2*2*{str_cm(data['I_U2_3'])}*{str_cm(data['Z_nagr3'])}/({str_cm(data['I1'])}/{str_cm(data['I2'])})",
            *form_bundle))
        worksheet.write_rich_string(f"C{row}", *get_rich_string(
            f"U|2MAX|^(3)^ = {str_cm(data['U2max3'])} В", *form_bundle))
        row += 1
    
        worksheet.write(f"A{row}", f"Условие обеспечения допустимого напряжения на выводах вторичной обмотки ТТ при 3-ф КЗ", form_def)        
        if data["U2max3"] < 1400:
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"U|2MAX|^(3)^ \u2264 1400 В;\n" +
                f"{str_cm(data['U2max3'])} В < 1400 В", *form_bundle))
            worksheet.write(f"C{row}", "Условие выполняется", form_def)
        else:
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"U|2MAX|^(3)^ \u2264 1400 В;\n" +
                f"{str_cm(data['U2max3'])} В > 1400 В", *form_bundle))
            worksheet.write(f"C{row}", "Условие не выполняется", form_bold)
        row += 1

    if (data["ignore_Int1"] == False) or (data["ignore_Ext1"] == False):
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"U|2MAX|^(1)^ - максимальное значение напряжения на выводах вторичной обмотки в режиме 1-ф КЗ, В",
            *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            f"U|2MAX|^(1)^ = \u221a2K|у|I|КЗ|Z|НАГР|/K|ТТ|;\n"
            "K|у| = 2;\n"
            f"U|2MAX|^(1)^ = \u221a2*2*{str_cm(data['I_U2_1'])}*{str_cm(data['Z_nagr1'])}/({str_cm(data['I1'])}/{str_cm(data['I2'])})",
            *form_bundle))
        worksheet.write_rich_string(f"C{row}", *get_rich_string(
            f"U|2MAX|^(1)^ = {str_cm(data['U2max1'])} В", *form_bundle))
        row += 1
        
        worksheet.write(f"A{row}", f"Условие обеспечения допустимого напряжения на выводах вторичной обмотки ТТ при 1-ф КЗ", form_def)        
        if data["U2max1"] < 1400:
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"U|2MAX|^(1)^ \u2264 1400 В;\n" +
                f"{str_cm(data['U2max1'])} В < 1400 В", *form_bundle))
            worksheet.write(f"C{row}", "Условие выполняется", form_def)
        else:
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"U|2MAX|^(1)^ \u2264 1400 В;\n" +
                f"{str_cm(data['U2max1'])} В > 1400 В", *form_bundle))
            worksheet.write(f"C{row}", "Условие не выполняется", form_bold)
        row += 1        

    #Проверка на 10%
    worksheet.merge_range(f"A{row}:C{row}", "Проверка ТТ на 10% погрешность", form_bold_center)
    row += 1

    #Проверка 10% 3-ф КЗ
    if data["ignore_K10_3"] == False:
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"K|ФАКТ|^(3)^ - фактический минимальный коэффициент предельной кратности при 3-ф КЗ",
            *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            f"K|ФАКТ|^(3)^ = K|НОМ|(R|2ТТ|+Z|НАГР.НОМ|)/ (R|2ТТ|+R|НАГР.РАСЧ|);\n"
            f"K|ФАКТ|^(3)^ = {str_cm(data['K_ras'])}*({str_cm(data['R2tt'])}+{str_cm(data['Z_ras'])})/({str_cm(data['R2tt'])}+{str_cm(data['R3'])})",
            *form_bundle))            
        worksheet.write_rich_string(f"C{row}", *get_rich_string(
            f"K|ФАКТ|^(3)^ = {str_cm(data['Kfact3'])}", *form_bundle))
        row += 1
        
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            "K|ПЕР| - коэффициент переходного режима",
            *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            f'K|ПЕР| = {str_cm(data["K10"])} [М.А.Беркович "Справочник по релейной защите", с.332]', *form_bundle))
        row += 1
    
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"K|ТРЕБ|^(3)^ - требуемый минимальный коэффициент предельной кратности при 3-ф КЗ",
            *form_bundle))
        if data["WM_ras"] == "Ext":
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"K|ТРЕБ|^(3)^ = K|ПЕР| I|РАСЧ|/I|1|;\n" +
                f"K|ТРЕБ|^(3)^ = {str_cm(data['K10'])}*{str_cm(data['I_ras_10_3'])}/{str_cm(data['I1'])};\n" +
                'I|РАСЧ| = I|КЗ ВНЕШ| - расчетный ток [М.А.Беркович "Справочник по релейной защите", с.332]',
                *form_bundle))
        elif data["WM_ras"] == "Int":
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"K|ТРЕБ|^(3)^ = K|ПЕР| I|РАСЧ|/I|1|;\n" +
                f"K|ТРЕБ|^(3)^ = {str_cm(data['K10'])}*{str_cm(data['I_ras_10_3'])}/{str_cm(data['I1'])};\n" +
                'I|РАСЧ| = I|КЗ ВНУТР| - расчетный ток [М.А.Беркович "Справочник по релейной защите", с.332]',
                *form_bundle))
        elif data["WM_ras"] == "Isz":
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"K|ТРЕБ|^(3)^ = K|ПЕР|I|РАСЧ|/I|1|;\n" + 
                f"K|ТРЕБ|^(3)^ = {str_cm(data['K10'])}*{str_cm(data['I_ras_10_3'])}/{str_cm(data['I1'])};\n" +
                'I|РАСЧ| = 1,1I|СРАБ.ЗАЩ| - расчетный ток [М.А.Беркович "Справочник по релейной защите", с.332]',
                *form_bundle))
        worksheet.write_rich_string(f"C{row}", *get_rich_string(
            f"K|ТРЕБ|^(3)^ = {str_cm(data['Ktreb3'])}", *form_bundle))
        row += 1
    
        worksheet.write(f"A{row}", "Условие правильной работы ТТ", form_def)
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            f"K|ТРЕБ|^(3)^ \u2264 K|ФАКТ|", *form_bundle))
        if data["Ktreb3"] < data["Kfact3"]:
            worksheet.write(f"C{row}", f"{str_cm(data['Ktreb3'])} < {str_cm(data['Kfact3'])};\n" +
                "Неравенство выполняется. ТТ подходит", form_def)
        else:
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"{str_cm(data['Ktreb3'])} > {str_cm(data['Kfact3'])};\n" +
                "Неравенство не выполняется. ?ТТ не подходит?", *form_bundle))
        row += 1

    #Проверка 10% 1-ф КЗ
    if data["ignore_K10_1"] == False:
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"K|ФАКТ|^(1)^ - фактический минимальный коэффициент предельной кратности при 1-ф КЗ",
            *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            f"K|ФАКТ|^(1)^ = K|НОМ|(R|2ТТ|+Z|НАГР.НОМ|)/ (R|2ТТ|+R|НАГР.РАСЧ|);\n"
            f"K|ФАКТ|^(1)^ = {str_cm(data['K_ras'])}*({str_cm(data['R2tt'])}+{str_cm(data['Z_ras'])})/({str_cm(data['R2tt'])}+{str_cm(data['R1'])})",
            *form_bundle))            
        worksheet.write_rich_string(f"C{row}", *get_rich_string(
            f"K|ФАКТ|^(1)^ = {str_cm(data['Kfact1'])}", *form_bundle))
        row += 1    
       
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"K|ТРЕБ|^(1)^ - требуемый минимальный коэффициент предельной кратности при 1-ф КЗ",
            *form_bundle))
        if data["WM_ras"] == "Ext":
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"K|ТРЕБ|^(1)^ = K|ПЕР| I|РАСЧ|/I|1|;\n" +
                f"K|ТРЕБ|^(1)^ = {str_cm(data['K10'])}*{str_cm(data['I_ras_10_1'])}/{str_cm(data['I1'])};\n" +
                'I|РАСЧ| = I|КЗ ВНЕШ| - расчетный ток [М.А.Беркович "Справочник по релейной защите", с.332]',
                *form_bundle))
        elif data["WM_ras"] == "Int":
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"K|ТРЕБ|^(1)^ = K|ПЕР| I|РАСЧ|/I|1|;\n" +
                f"K|ТРЕБ|^(1)^ = {str_cm(data['K10'])}*{str_cm(data['I_ras_10_1'])}/{str_cm(data['I1'])};\n" +
                'I|РАСЧ| = I|КЗ ВНТУР| - расчетный ток [М.А.Беркович "Справочник по релейной защите", с.332]',
                *form_bundle))
        elif data["WM_ras"] == "Isz":
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"K|ТРЕБ|^(1)^ = K|ПЕР|I|РАСЧ|/I|1|;\n" + 
                f"K|ТРЕБ|^(1)^ = {str_cm(data['K10'])}*{str_cm(data['I_ras_10_1'])}/{str_cm(data['I1'])};\n" +
                'I|РАСЧ| = 1,1I|СРАБ.ЗАЩ| - расчетный ток [М.А.Беркович "Справочник по релейной защите", с.332]',
                *form_bundle))
        worksheet.write_rich_string(f"C{row}", *get_rich_string(
            f"K|ТРЕБ|^(1)^ = {str_cm(data['Ktreb1'])}", *form_bundle))
        row += 1
        
        worksheet.write(f"A{row}", "Условие правильной работы ТТ", form_def)
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            f"K|ТРЕБ|^(1)^ \u2264 K|ФАКТ|", *form_bundle))
        if data["Ktreb1"] < data["Kfact1"]:
            worksheet.write(f"C{row}", f"{str_cm(data['Ktreb1'])} < {str_cm(data['Kfact1'])};\n" +
                "Неравенство выполняется. ТТ подходит", form_def)
        else:
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"{str_cm(data['Ktreb1'])} > {str_cm(data['Kfact1'])};\n" +
                "Неравенство не выполняется. ?ТТ не подходит?", *form_bundle))
        row += 1

    #Проверка 10% 2-ф КЗ
    if data["calc10_2ph"] == True:
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"K|ФАКТ|^(2)^ - фактический минимальный коэффициент предельной кратности при 2-ф КЗ",
            *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            f"K|ФАКТ|^(2)^ = K|НОМ|(R|2ТТ|+Z|НАГР.НОМ|)/ (R|2ТТ|+R|НАГР.РАСЧ|);\n"
            f"K|ФАКТ|^(2)^ = {str_cm(data['K_ras'])}*({str_cm(data['R2tt'])}+{str_cm(data['Z_ras'])})/({str_cm(data['R2tt'])}+{str_cm(data['R2'])})",
            *form_bundle))            
        worksheet.write_rich_string(f"C{row}", *get_rich_string(
            f"K|ФАКТ|^(2)^ = {str_cm(data['Kfact2'])}", *form_bundle))
        row += 1
        
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            "K|ПЕР| - коэффициент переходного режима",
            *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            f'K|ПЕР| = {str_cm(data["K10"])} [Чернобровов Н.В. "Релейная защита", с.98]', *form_bundle))
        row += 1
    
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            f"K|ТРЕБ|^(2)^ - требуемый минимальный коэффициент предельной кратности при 2-ф КЗ",
            *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            f"K|ТРЕБ|^(2)^ = K|ПЕР|I|РАСЧ|/I|1|;\n" + 
            f"K|ТРЕБ|^(2)^ = {str_cm(data['K10'])}*{str_cm(data['I_ras_10_3'])}/{str_cm(data['I1'])};\n" +
            'I|РАСЧ| = 1,1I|СРАБ.ЗАЩ| - расчетный ток [М.А.Беркович "Справочник по релейной защите", с.332]',
            *form_bundle))
        worksheet.write_rich_string(f"C{row}", *get_rich_string(
            f"K|ТРЕБ|^(2)^ = {str_cm(data['Ktreb3'])}", *form_bundle))
        row += 1
    
        worksheet.write(f"A{row}", "Условие правильной работы ТТ", form_def)
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            f"K|ТРЕБ|^(2)^ \u2264 K|ФАКТ|", *form_bundle))
        if data["Ktreb3"] < data["Kfact2"]:
            worksheet.write(f"C{row}", f"{str_cm(data['Ktreb3'])} < {str_cm(data['Kfact2'])};\n" +
                "Неравенство выполняется. ТТ подходит", form_def)
        else:
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"{str_cm(data['Ktreb3'])} > {str_cm(data['Kfact2'])};\n" +
                "Неравенство не выполняется. ?ТТ не подходит?", *form_bundle))
        row += 1

    #Определение времени до насыщения
    worksheet.merge_range(f"A{row}:C{row}", "Определение времени до насыщения", form_bold_center)
    row += 1
        
    worksheet.write_rich_string(f"A{row}", *get_rich_string(
        "z|2НОМ| - номинальная вторичная нагрузка ТТ, Ом",
        *form_bundle))
    worksheet.write_rich_string(f"B{row}", *get_rich_string(
        "z|2НОМ| = \u221a((R|2ТТ|+z|НАГР.НОМ| cos\u03c6|НОМ|)^2^+\n" +
        "(x|2ТТ| + z|НАГР.НОМ| sin\u03c6|НОМ|)^2^);\n" +
        f"z|2НОМ| = \u221a(({str_cm(data['R2tt'])}+{str_cm(data['Z_ras'])}*{str_cm(data['cosf_tt'])})^2^"+
        f"+({str_cm(data['X2tt'])}+{str_cm(data['Z_ras'])}*{str_cm(data['sinf_tt'])})^2^)",
        *form_bundle))
    worksheet.write_rich_string(f"C{row}", *get_rich_string(
        f"z|2НОМ| = {str_cm(data['z2_ras'])} Ом", *form_bundle))
    row += 1

    #Время до насыщения при 3-ф КЗ
    if (data["ignore_Int3"] == False) or (data["ignore_Ext3"] == False):
        worksheet.merge_range(f"A{row}:C{row}", "Определение времени до насыщения при 3-ф КЗ", form_bold_center)
        row += 1
        
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            "z|2РАСЧ| - расчетная вторичная нагрузка ТТ при 3-ф КЗ, Ом",
            *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            "z|2РАСЧ| = \u221a((R|2ТТ|+R|НАГР|)^2^+(x|2ТТ|+x|НАГР|)^2^);\n" +
            f"z|2РАСЧ| = \u221a(({str_cm(data['R2tt'])}+{str_cm(data['R3'])})^2^"
            f"+({str_cm(data['X2tt'])}+{str_cm(data['X3'])})^2^)",
            *form_bundle))
        worksheet.write_rich_string(f"C{row}", *get_rich_string(
            f"z|2РАСЧ| = {str_cm(data['z2_fact3'])} Ом", *form_bundle))
        row += 1
        
        #Без остаточной намагниченности внутреннее 3-ф
        if data["ignore_Int3"] == False:
            worksheet.merge_range(f"A{row}:C{row}", "Определение tнас без учета остаточной намагниченности при внутр. 3-ф КЗ",
                                  form_bold_center)
            row += 1
            
            worksheet.write(f"A{row}", "A - коэффициент, учитывающий соотношение между номинальными параметрами ТТ и реальными в месте его установки", form_def)
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                "A = I|1| K|НОМ| Z|2НОМ|/(I|КЗ| Z|2РАСЧ|);\n" +
                f"A = {str_cm(data['I1'])}*{str_cm(data['K_ras'])}*{str_cm(data['z2_ras'])}/({str_cm(data['Ikz3_Int'])}*{str_cm(data['z2_fact3'])})",
                *form_bundle))
            worksheet.write(f"C{row}", f"A = {str_cm(data['A3_Int_0'])}", form_def)
            row += 1
            
            worksheet.set_row(row-1, 225)
            worksheet.merge_range(f"A{row}:B{row}", "", form_def)
            worksheet.insert_image(f"A{row}", "Graph_Int3_gamma_0.png", {"x_scale": 0.7, "y_scale":0.7, "align":"center"})
            if data["t_nas3_Int_0"] > data["t_nas_treb_Int"]:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas3_Int_0'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas3_Int_0'])} мс > {str_cm(data['t_nas_treb_Int'])} мс;\n" +
                    "Условие выполняется", *form_bundle))
            elif data["t_nas3_Int_0"] == 0:
                worksheet.write(f"C{row}", "ТТ не насыщается", form_def)
            else:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas3_Int_0'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas3_Int_0'])} мс < {str_cm(data['t_nas_treb_Int'])} мс;\n" +
                    "?Условие не выполняется?", *form_bundle))
            row += 1

            #С остаточной намагниченностью внутренее 3-ф
            worksheet.merge_range(f"A{row}:C{row}", "Определение tнас с учетом остаточной намагниченности при внутр. 3-ф КЗ",
                                  form_bold_center)
            row += 1
            
            worksheet.write(f"A{row}", "A - коэффициент, учитывающий соотношение между номинальными параметрами ТТ и реальными в месте его установки", form_def)
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"A = {str_cm(data['A3_Int_0'])};\n" +
                f"A(1-K|\u03b3|) = {str_cm(data['A3_Int_0'])}*(1-{str_cm(data['k_gamma'])})",
                *form_bundle))
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"A(1-K|\u03b3|) = {str_cm(data['A3_Int'])}", *form_bundle))
            row += 1
            
            worksheet.set_row(row-1, 225)
            worksheet.merge_range(f"A{row}:B{row}", "", form_def)
            worksheet.insert_image(f"A{row}", "Graph_Int3_gamma_" + str(data["k_gamma"]) + ".png", {"x_scale": 0.7, "y_scale":0.7})
            if data["t_nas3_Int"] > data["t_nas_treb_Int"]:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas3_Int'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas3_Int'])} мс > {str_cm(data['t_nas_treb_Int'])} мс;\n" +
                    "Условие выполняется", *form_bundle))
            elif data["t_nas3_Int"] == 0:
                worksheet.write(f"C{row}", "ТТ не насыщается", form_def)
            else:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas3_Int'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas3_Int'])} мс < {str_cm(data['t_nas_treb_Int'])} мс;\n" +
                    "?Условие не выполняется?", *form_bundle))
            row += 1    

        #Без остаточной намагниченности внешнее 3-ф
        if data["ignore_Ext3"] == False:
            worksheet.merge_range(f"A{row}:C{row}", "Определение tнас без учета остаточной намагниченности при 3-ф внеш. КЗ",
                                  form_bold_center)
            row += 1
            
            worksheet.write(f"A{row}", "A - коэффициент, учитывающий соотношение между номинальными параметрами ТТ и реальными в месте его установки", form_def)
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                "A = I|1| K|НОМ| Z|2НОМ|/(I|КЗ| Z|2РАСЧ|);\n" +
                f"A = {str_cm(data['I1'])}*{str_cm(data['K_ras'])}*{str_cm(data['z2_ras'])}/({str_cm(data['Ikz3_Ext'])}*{str_cm(data['z2_fact3'])})",
                *form_bundle))
            worksheet.write(f"C{row}", f"A = {str_cm(data['A3_Ext_0'])}", form_def)
            row += 1
            
            worksheet.merge_range(f"A{row}:B{row}", "", form_def)
            worksheet.set_row(row-1, 225)
            worksheet.insert_image(f"A{row}", "Graph_Ext3_gamma_0.png", {"x_scale": 0.7, "y_scale":0.7})
            if data["t_nas3_Ext_0"] > data["t_nas_treb_Ext"]:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas3_Ext_0'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas3_Ext_0'])} мс > {str_cm(data['t_nas_treb_Ext'])} мс;\n" +
                    "Условие выполняется", *form_bundle))
            elif data["t_nas3_Ext_0"] == 0:
                worksheet.write(f"C{row}", "ТТ не насыщается", form_def)
            else:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas3_Ext_0'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas3_Ext_0'])} мс < {str_cm(data['t_nas_treb_Ext'])} мс;\n" +
                    "?Условие не выполняется?", *form_bundle))
            row += 1

            #С остаточной намагниченностью внешнее 3-ф
            worksheet.merge_range(f"A{row}:C{row}", "Определение tнас с учетом остаточной намагниченности при внеш. 3-ф КЗ",
                                  form_bold_center)
            row += 1
            
            worksheet.write(f"A{row}", "A - коэффициент, учитывающий соотношение между номинальными параметрами ТТ и реальными в месте его установки", form_def)
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"A = {str_cm(data['A3_Ext_0'])};\n" +
                f"A(1-K|\u03b3|) = {str_cm(data['A3_Ext_0'])}*(1-{str_cm(data['k_gamma'])})",
                *form_bundle))
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"A(1-K|\u03b3|) = {str_cm(data['A3_Ext'])}", *form_bundle))
            row += 1
            
            worksheet.set_row(row-1, 225)
            worksheet.merge_range(f"A{row}:B{row}", "", form_def)
            worksheet.insert_image(f"A{row}", "Graph_Ext3_gamma_" + str(data["k_gamma"]) + ".png", {"x_scale": 0.7, "y_scale":0.7})
            if data["t_nas3_Ext"] > data["t_nas_treb_Ext"]:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas3_Ext'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas3_Ext'])} мс > {str_cm(data['t_nas_treb_Ext'])} мс;\n" +
                    "Условие выполняется", *form_bundle))
            elif data["t_nas3_Ext"] == 0:
                worksheet.write(f"C{row}", "ТТ не насыщается", form_def)
            else:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas3_Ext'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas3_Ext'])} мс < {str_cm(data['t_nas_treb_Ext'])} мс;\n" +
                    "?Условие не выполняется?", *form_bundle))
            row += 1

    #Время до насыщения при 1-ф КЗ
    if (data["ignore_Int1"] == False) or (data["ignore_Ext1"] == False):
        worksheet.merge_range(f"A{row}:C{row}", "Определение времени до насыщения при 1-ф КЗ", form_bold_center)
        row += 1
          
        worksheet.write_rich_string(f"A{row}", *get_rich_string(
            "z|2РАСЧ| - расчетная вторичная нагрузка ТТ при 1-ф КЗ, Ом",
            *form_bundle))
        worksheet.write_rich_string(f"B{row}", *get_rich_string(
            "z|2РАСЧ| = \u221a((R|2ТТ|+R|НАГР|)^2^+(x|2ТТ|+x|НАГР|)^2^);\n" +
            f"z|2РАСЧ| = \u221a(({str_cm(data['R2tt'])}+{str_cm(data['R1'])})^2^"
            f"+({str_cm(data['X2tt'])}+{str_cm(data['X1'])})^2^)",
            *form_bundle))
        worksheet.write_rich_string(f"C{row}", *get_rich_string(
            f"z|2РАСЧ| = {str_cm(data['z2_fact1'])} Ом", *form_bundle))
        row += 1
        
        #Без остаточной намагниченности внутреннее 1-ф
        if data["ignore_Int1"] == False:
            worksheet.merge_range(f"A{row}:C{row}", "Определение tнас без учета остаточной намагниченности при внутр. 1-ф КЗ",
                                  form_bold_center)
            row += 1
            
            worksheet.write(f"A{row}", "A - коэффициент, учитывающий соотношение между номинальными параметрами ТТ и реальными в месте его установки", form_def)
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                "A = I|1| K|НОМ| Z|2НОМ|/(I|КЗ| Z|2РАСЧ|);\n" +
                f"A = {str_cm(data['I1'])}*{str_cm(data['K_ras'])}*{str_cm(data['z2_ras'])}/({str_cm(data['Ikz1_Int'])}*{str_cm(data['z2_fact1'])})",
                *form_bundle))
            worksheet.write(f"C{row}", f"A = {str_cm(data['A1_Int_0'])}", form_def)
            row += 1
          
            worksheet.set_row(row-1, 225)
            worksheet.merge_range(f"A{row}:B{row}", "", form_def)
            worksheet.insert_image(f"A{row}", "Graph_Int1_gamma_0.png", {"x_scale": 0.7, "y_scale":0.7, "align":"center"})
            if data["t_nas1_Int_0"] > data["t_nas_treb_Int"]:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas1_Int_0'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas1_Int_0'])} мс > {str_cm(data['t_nas_treb_Int'])} мс;\n" +
                    "Условие выполняется", *form_bundle))
            elif data["t_nas1_Int_0"] == 0:
                worksheet.write(f"C{row}", "ТТ не насыщается", form_def)
            else:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas1_Int_0'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas1_Int_0'])} мс < {str_cm(data['t_nas_treb_Int'])} мс;\n" +
                    "?Условие не выполняется?", *form_bundle))
            row += 1

            #С остаточной намагниченностью внутренее 1-ф
            worksheet.merge_range(f"A{row}:C{row}", "Определение tнас с учетом остаточной намагниченности при внутр. 1-ф КЗ",
                                  form_bold_center)
            row += 1
            
            worksheet.write(f"A{row}", "A - коэффициент, учитывающий соотношение между номинальными параметрами ТТ и реальными в месте его установки", form_def)
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"A = {str_cm(data['A1_Int_0'])};\n" +
                f"A(1-K|\u03b3|) = {str_cm(data['A1_Int_0'])}*(1-{str_cm(data['k_gamma'])})",
                *form_bundle))
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"A(1-K|\u03b3|) = {str_cm(data['A1_Int'])}", *form_bundle))
            row += 1
            
            worksheet.set_row(row-1, 225)
            worksheet.merge_range(f"A{row}:B{row}", "", form_def)
            worksheet.insert_image(f"A{row}", "Graph_Int1_gamma_" + str(data["k_gamma"]) + ".png", {"x_scale": 0.7, "y_scale":0.7})
            if data["t_nas1_Int"] > data["t_nas_treb_Int"]:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas1_Int'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas1_Int'])} мс > {str_cm(data['t_nas_treb_Int'])} мс;\n" +
                    "Условие выполняется", *form_bundle))
            elif data["t_nas1_Int"] == 0:
                worksheet.write(f"C{row}", "ТТ не насыщается", form_def)
            else:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas1_Int'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas1_Int'])} мс < {str_cm(data['t_nas_treb_Int'])} мс;\n" +
                    "?Условие не выполняется?", *form_bundle))
            row += 1

        #Без остаточной намагниченности внешнее 1-ф
        if data["ignore_Ext1"] == False:
            worksheet.merge_range(f"A{row}:C{row}", "Определение tнас без учета остаточной намагниченности при 1-ф внеш. КЗ",
                                  form_bold_center)
            row += 1
            
            worksheet.write(f"A{row}", "A - коэффициент, учитывающий соотношение между номинальными параметрами ТТ и реальными в месте его установки", form_def)
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                "A = I|1| K|НОМ| Z|2НОМ|/(I|КЗ| Z|2РАСЧ|);\n" +
                f"A = {str_cm(data['I1'])}*{str_cm(data['K_ras'])}*{str_cm(data['z2_ras'])}/({str_cm(data['Ikz1_Ext'])}*{str_cm(data['z2_fact1'])})",
                *form_bundle))
            worksheet.write(f"C{row}", f"A = {str_cm(data['A1_Ext_0'])}", form_def)
            row += 1
            
            worksheet.merge_range(f"A{row}:B{row}", "", form_def)
            worksheet.set_row(row-1, 225)
            worksheet.insert_image(f"A{row}", "Graph_Ext1_gamma_0.png", {"x_scale": 0.7, "y_scale":0.7})
            if data["t_nas1_Ext_0"] > data["t_nas_treb_Ext"]:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas1_Ext_0'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas1_Ext_0'])} мс > {str_cm(data['t_nas_treb_Ext'])} мс;\n" +
                    "Условие выполняется", *form_bundle))
            elif data["t_nas1_Ext_0"] == 0:
                worksheet.write(f"C{row}", "ТТ не насыщается", form_def)
            else:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas1_Ext_0'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas1_Ext_0'])} мс < {str_cm(data['t_nas_treb_Ext'])} мс;\n" +
                    "?Условие не выполняется?", *form_bundle))
            row += 1

            #С остаточной намагниченностью внешнее 1-ф
            worksheet.merge_range(f"A{row}:C{row}", "Определение tнас с учетом остаточной намагниченности при внеш. 1-ф КЗ",
                                  form_bold_center)
            row += 1
            
            worksheet.write(f"A{row}", "A - коэффициент, учитывающий соотношение между номинальными параметрами ТТ и реальными в месте его установки", form_def)
            worksheet.write_rich_string(f"B{row}", *get_rich_string(
                f"A = {str_cm(data['A1_Ext_0'])};\n" +
                f"A(1-K|\u03b3|) = {str_cm(data['A1_Ext_0'])}*(1-{str_cm(data['k_gamma'])})",
                *form_bundle))
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                f"A(1-K|\u03b3|) = {str_cm(data['A1_Ext'])}", *form_bundle))
            row += 1
            
            worksheet.set_row(row-1, 225)
            worksheet.merge_range(f"A{row}:B{row}", "", form_def)
            worksheet.insert_image(f"A{row}", "Graph_Ext1_gamma_" + str(data["k_gamma"]) + ".png", {"x_scale": 0.7, "y_scale":0.7})
            if data["t_nas1_Ext"] > data["t_nas_treb_Ext"]:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas1_Ext'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas1_Ext'])} мс > {str_cm(data['t_nas_treb_Ext'])} мс;\n" +
                    "Условие выполняется", *form_bundle))
            elif data["t_nas1_Ext"] == 0:
                worksheet.write(f"C{row}", "ТТ не насыщается", form_def)
            else:
                worksheet.write_rich_string(f"C{row}", *get_rich_string(
                    f"t|НАС| = {str_cm(data['t_nas1_Ext'])} мс;\n" +
                    "\n" +
                    f"{str_cm(data['t_nas1_Ext'])} мс < {str_cm(data['t_nas_treb_Ext'])} мс;\n" +
                    "?Условие не выполняется?", *form_bundle))
                
    try:
        workbook.close()
    except xlsxwriter.exceptions.FileCreateError as error:
        messagebox.showerror(message="Close the Excel file first!")
        return
    
    startfile("log_fixed.xlsx")

def save_report_short (data, italic):

    workbook = xlsxwriter.Workbook("log_short.xlsx")
    worksheet = workbook.add_worksheet()

    form_super = workbook.add_format({"italic": italic, "font_script": 1, "font_name": "ISOCPEUR", "font_size": 16})
    form_super_bold = workbook.add_format({"italic": italic, "bold":1, "font_script": 1, "font_name": "ISOCPEUR", "font_size": 16})
    form_sub = workbook.add_format({"italic": italic, "font_script": 2, "font_name": "ISOCPEUR", "font_size": 16})
    form_def = workbook.add_format({"italic": italic, "valign": "vcenter", "font_name": "ISOCPEUR", "font_size": 12, "text_wrap":1, "border": 1})
    form_def_rotated = workbook.add_format({"italic": italic, "valign": "vcenter", "rotation": 90, "font_name": "ISOCPEUR", "font_size": 12, "text_wrap":1, "border": 1})
    form_bold = workbook.add_format({"italic": italic, "valign": "vcenter", "bold": 1, "font_name": "ISOCPEUR", "font_size": 12, "border": 1})
    form_bold_center = workbook.add_format({"italic": italic, "bold": 1, "align": "center", "valign": "vcenter", "font_name": "ISOCPEUR", "font_size": 12, "border": 1})
    form_bundle = [form_sub, form_super, form_def, form_bold]
    
    worksheet.set_column(0, 0, 10)
    worksheet.set_column(1, 1, 40)
    worksheet.set_column(2, 4, 25)

    row = 1

    #Параметры ТТ
    row_begin = row
    
    worksheet.write(f"B{row}", "Наименование присоединения", form_def)
    row += 1
    
    worksheet.write_rich_string(f"B{row}", *get_rich_string("I|1| - номинальный первичный ток, А", *form_bundle))
    worksheet.write(f"C{row}", str_cm(data['I1']), form_def)
    row += 1

    worksheet.write_rich_string(f"B{row}", *get_rich_string("I|2| - номинальный вторичный ток, А", *form_bundle))
    worksheet.write(f"C{row}", str_cm(data['I2']), form_def)
    row += 1

    worksheet.write(f"B{row}", "Класс точности", form_def)
    if data["k_gamma"] > 0.1:
        worksheet.write(f"C{row}", "10P", form_def)
    else:
        worksheet.write(f"C{row}", "10PR", form_def)
    row += 1

    worksheet.write_rich_string(f"B{row}", *get_rich_string("S|НОМ| - номинальная нагрузка, ВА", *form_bundle))
    worksheet.write(f"C{row}", str_cm(data['S_ras']), form_def)
    row += 1

    worksheet.write_rich_string(f"B{row}", *get_rich_string("K|НОМ| - номинальная предельная кратность", *form_bundle))
    worksheet.write(f"C{row}", str_cm(data['K_ras']), form_def)
    row += 1

    worksheet.write_rich_string(f"B{row}", *get_rich_string("R|2| - активное сопротивление вторичной обмотки, Ом", *form_bundle))
    worksheet.write(f"C{row}", str_cm(data['R2tt']), form_def)
    row += 1

    if data["X2tt"] != 0:
        worksheet.write_rich_string(f"B{row}", *get_rich_string("X|2| - реактивное сопротивление вторичной обмотки, Ом", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['X2tt']), form_def)
        row += 1

    worksheet.merge_range(f"A{row_begin}:A{row-1}", "Паспортные данные", form_def_rotated)

    #Значение для 3-ф КЗ
    row_begin = row
    
    if data["ignore_Int3"] == False:
        worksheet.write_rich_string(f"B{row}", *get_rich_string("I|ВНУТР|^(3)^ - ток 3-ф КЗ в зоне действия защиты, А", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Ikz3_Int']), form_def)
        row += 1

        worksheet.write_rich_string(f"B{row}", *get_rich_string("T|P ВНУТР|^(3)^ - постоянная времени апериод. сост. 3-ф тока КЗ в зоне действия защиты, мс", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Tp3_Int']), form_def)
        row += 1

    if data["ignore_Ext3"] == False:
        worksheet.write_rich_string(f"B{row}", *get_rich_string("I|ВНЕШ|^(3)^ - ток 3-ф КЗ вне зоны действия защиты, А", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Ikz3_Ext']), form_def)
        row += 1

        worksheet.write_rich_string(f"B{row}", *get_rich_string("T|P ВНЕШ|^(3)^ - постоянная времени апериод. сост. 3-ф тока КЗ вне зоны действия защиты, мс", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Tp3_Ext']), form_def)
        row += 1

    #Значения для 1-ф КЗ
    if data["ignore_Int1"] == False:
        worksheet.write_rich_string(f"B{row}", *get_rich_string("I|ВНУТР|^(1)^ - ток 1-ф КЗ в зоне действия защиты, А", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Ikz1_Int']), form_def)
        row += 1

        worksheet.write_rich_string(f"B{row}", *get_rich_string("T|P ВНУТР|^(1)^ - постоянная времени апериод. сост. 1-ф тока КЗ в зоне действия защиты, мс", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Tp1_Int']), form_def)
        row += 1

    if data["ignore_Ext1"] == False:
        worksheet.write_rich_string(f"B{row}", *get_rich_string("I|ВНЕШ|^(1)^ - ток 1-ф КЗ вне зоны действия защиты, А", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Ikz1_Ext']), form_def)
        row += 1

        worksheet.write_rich_string(f"B{row}", *get_rich_string("T|P ВНЕШ|^(1)^ - постоянная времени апериод. сост. 1-ф тока КЗ вне зоны действия защиты, мс", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Tp1_Ext']), form_def)
        row += 1

    worksheet.merge_range(f"A{row_begin}:A{row-1}", "Параметры расчетных режимов", form_def_rotated)

    #Параметры нагрузки
    row_begin = row
    
    worksheet.write(f"B{row}", "Схема соединения ТТ", form_def)
    if data["CON_type"] == "star":
        worksheet.write(f"C{row}", "Звезда", form_def)
    elif data["CON_type"] == "triangle":
        worksheet.write(f"C{row}", "Треугольник", form_def)
    elif data["CON_type"] == "part Y":
        worksheet.write(f"C{row}", "Неполная звезда", form_def)
    row += 1

    worksheet.write_rich_string(f"B{row}", *get_rich_string("S|КАБ| - поперечное сечение контрольного кабеля, мм^2^", *form_bundle))
    worksheet.write(f"C{row}", str_cm(data['Pop_ras']), form_def)
    row += 1

    worksheet.write_rich_string(f"B{row}", *get_rich_string("S|НАГР| - нагрузка фазы ТТ, ВА", *form_bundle))
    worksheet.write(f"C{row}", str_cm(data['S_rele']), form_def)
    row += 1

    worksheet.write_rich_string(f"B{row}", *get_rich_string("cos\u03c6|НАГР| - коэффициент мощности нагрузки", *form_bundle))
    worksheet.write(f"C{row}", str_cm(data['cosf_rele']), form_def)
    row += 1

    if (data["ignore_Int3"] == False) or (data["ignore_Ext3"] == False):
        if data["X3"] == 0:
            worksheet.write_rich_string(f"B{row}", *get_rich_string("R|НАГР| - cуммарное сопротивление нагрузки при 3-ф КЗ, Ом", *form_bundle))
            worksheet.write(f"C{row}", str_cm(data['Z_nagr3']), form_def)
        else:
            worksheet.write(f"B{row}", "Суммарное сопротивление нагрузки при 3-ф КЗ, Ом", form_def)
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                            f"R|НАГР| = {str_cm(data['R3'])};\n" +
                            f"X|НАГР| = {str_cm(data['X3'])};\n" +
                            f"Z|НАГР| = {str_cm(data['Z_nagr3'])}",
                            *form_bundle))
        row += 1

    if (data["ignore_Int1"] == False) or (data["ignore_Ext1"] == False):
        if data["X1"] == 0:
            worksheet.write_rich_string(f"B{row}", *get_rich_string("R|НАГР| - cуммарное сопротивление нагрузки при 1-ф КЗ, Ом", *form_bundle))
            worksheet.write(f"C{row}", str_cm(data['Z_nagr1']), form_def)
        else:
            worksheet.write(f"B{row}", "Суммарное сопротивление нагрузки при 1-ф КЗ, Ом", form_def)
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                            f"R|НАГР| = {str_cm(data['R1'])};\n" +
                            f"X|НАГР| = {str_cm(data['X1'])};\n" +
                            f"Z|НАГР| = {str_cm(data['Z_nagr1'])}",
                            *form_bundle))
        row += 1

    if data["calc10_2ph"] == True:
        if data["X3"] == 0:
            worksheet.write_rich_string(f"B{row}", *get_rich_string("R|НАГР| - cуммарное сопротивление нагрузки при 2-ф КЗ, Ом", *form_bundle))
            worksheet.write(f"C{row}", str_cm(data['Z_nagr2']), form_def)
        else:
            worksheet.write(f"B{row}", "Суммарное сопротивление нагрузки при 2-ф КЗ, Ом", form_def)
            worksheet.write_rich_string(f"C{row}", *get_rich_string(
                            f"R|НАГР| = {str_cm(data['R2'])};\n" +
                            f"X|НАГР| = {str_cm(data['X3'])};\n" +
                            f"Z|НАГР| = {str_cm(data['Z_nagr2'])}",
                            *form_bundle))
        row += 1
        

    worksheet.merge_range(f"A{row_begin}:A{row-1}", "Параметры нагрузки ТТ", form_def_rotated)

    #Проверка по U2
    row_begin = row
    
    if (data["ignore_Int3"] == False) or (data["ignore_Ext3"] == False):
        worksheet.write_rich_string(f"B{row}", *get_rich_string("U|2MAX|^(3)^ - напряжение на вторичной обмотке при 3-ф КЗ", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['U2max3']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Допустимое напряжение на вторичной обмотке ТТ при 3-ф КЗ, В", form_def)
        if data["U2max3"] < 1400:
            worksheet.write(f"C{row}", f"{str_cm(data['U2max3'])} < 1400", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['U2max3'])} > 1400", form_bold)
        row += 1        

    if (data["ignore_Int1"] == False) or (data["ignore_Ext1"] == False):
        worksheet.write_rich_string(f"B{row}", *get_rich_string("U|2MAX|^(1)^ - напряжение на вторичной обмотке при 1-ф КЗ", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['U2max1']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Допустимое напряжение на вторичной обмотке ТТ при 1-ф КЗ, В", form_def)
        if data["U2max1"] < 1400:
            worksheet.write(f"C{row}", f"{str_cm(data['U2max1'])} < 1400", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['U2max1'])} > 1400", form_bold)
        row += 1

    worksheet.merge_range(f"A{row_begin}:A{row-1}", "Проверка наряжения на вторичной обмотке", form_def_rotated)

    #Проверка 10%
    row_begin = row
    
    if data["WM_ras"] == "Isz":
        worksheet.write_rich_string(f"B{row}", *get_rich_string("I|РАСЧ| - расчетный ток для расчета 10% погрешности, А", *form_bundle))
        worksheet.write_rich_string(f"C{row}", *get_rich_string(f"I|РАСЧ| = 1.1I|СЗ| = {str_cm(data['I_ras_10_3'])}", *form_bundle))
        row += 1
        
    if data["ignore_K10_3"] == False:
        if data["WM_ras"] == "Ext":
            worksheet.write_rich_string(f"B{row}", *get_rich_string("I|РАСЧ| - расчетный ток для расчета 10% погрешности при 3-ф КЗ, А", *form_bundle))
            worksheet.write_rich_string(f"C{row}", *get_rich_string(f"I|РАСЧ| = I|КЗ ВНЕШ|^(3)^ = {str_cm(data['I_ras_10_3'])}", *form_bundle))
            row += 1
            
        worksheet.write_rich_string(f"B{row}", *get_rich_string("K|ФАКТ|^(3)^ - фактический коэффициент предельной кратности при 3-ф КЗ", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Kfact3']), form_def)
        row += 1

        worksheet.write_rich_string(f"B{row}", *get_rich_string("K|ТРЕБ|^(3)^ - требуемый коэффициент предельной кратности при 3-ф КЗ", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Ktreb3']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Условие правильной работы при 3-ф КЗ", form_def)
        if data["Kfact3"] > data["Ktreb3"]:
            worksheet.write(f"C{row}", f"{str_cm(data['Kfact3'])} > {str_cm(data['Ktreb3'])}", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['Kfact3'])} < {str_cm(data['Ktreb3'])}", form_bold)
        row += 1

    if data["ignore_K10_1"] == False:
        if data["WM_ras"] == "Ext":
            worksheet.write_rich_string(f"B{row}", *get_rich_string("I|РАСЧ| - расчетный ток для расчета 10% погрешности при 1-ф КЗ, А", *form_bundle))
            worksheet.write_rich_string(f"C{row}", *get_rich_string(f"I|РАСЧ| = I|КЗ ВНЕШ|^(1)^ = {str_cm(data['I_ras_10_1'])}", *form_bundle))
            row += 1
            
        worksheet.write_rich_string(f"B{row}", *get_rich_string("K|ФАКТ|^(1)^ - фактический коэффициент предельной кратности при 1-ф КЗ", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Kfact1']), form_def)
        row += 1

        worksheet.write_rich_string(f"B{row}", *get_rich_string("K|ТРЕБ|^(1)^ - требуемый коэффициент предельной кратности при 1-ф КЗ", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Ktreb1']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Условие правильной работы при 1-ф КЗ", form_def)
        if data["Kfact1"] > data["Ktreb1"]:
            worksheet.write(f"C{row}", f"{str_cm(data['Kfact1'])} > {str_cm(data['Ktreb1'])}", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['Kfact1'])} < {str_cm(data['Ktreb1'])}", form_bold)
        row += 1

    if data["calc10_2ph"] == True:
           
        worksheet.write_rich_string(f"B{row}", *get_rich_string("K|ФАКТ|^(2)^ - фактический коэффициент предельной кратности при 2-ф КЗ", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Kfact2']), form_def)
        row += 1

        worksheet.write_rich_string(f"B{row}", *get_rich_string("K|ТРЕБ|^(2)^ - требуемый коэффициент предельной кратности при 2-ф КЗ", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['Ktreb3']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Условие правильной работы при 2-ф КЗ", form_def)
        if data["Kfact2"] > data["Ktreb3"]:
            worksheet.write(f"C{row}", f"{str_cm(data['Kfact2'])} > {str_cm(data['Ktreb3'])}", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['Kfact2'])} < {str_cm(data['Ktreb3'])}", form_bold)
        row += 1

    worksheet.merge_range(f"A{row_begin}:A{row-1}", "Проверка на 10% погрешность", form_def_rotated)

    #Насыщение
    row_begin = row
    
    worksheet.write_rich_string(f"B{row}", *get_rich_string("K|\u03b3| - коэффициент остаточной намагниченности", *form_bundle))
    worksheet.write(f"C{row}", str_cm(data['k_gamma']), form_def)
    row += 1

    if data["t_nas_treb_Int"] == data["t_nas_treb_Ext"]:
        worksheet.write(f"B{row}", "Требуемое время до насыщения, мс", form_def)
        worksheet.write(f"C{row}", str_cm(data['t_nas_treb_Int']), form_def)
        row += 1
    else:
        worksheet.write(f"B{row}", "Требуемое время до насыщения при КЗ в зоне дейтсвия защиты, мс", form_def)
        worksheet.write(f"C{row}", str_cm(data['t_nas_treb_Int']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Требуемое время до насыщения при КЗ вне зоны дейтсвия защиты, мс", form_def)
        worksheet.write(f"C{row}", str_cm(data['t_nas_treb_Ext']), form_def)
        row += 1
    
    #3-ф в зоне действия защиты
    if data["ignore_Int3"] == False:
        worksheet.write(f"B{row}", "A - значение коэффициента при 3-ф КЗ в зоне действия защиты", form_def)
        worksheet.write(f"C{row}", str_cm(data['A3_Int_0']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Время до насыщения при 3-ф КЗ в зоне действия защиты, мс", form_def)
        if data["t_nas3_Int_0"] > data["t_nas_treb_Int"]:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas3_Int_0'])} > {str_cm(data['t_nas_treb_Int'])}", form_def)
        elif data["t_nas3_Int_0"] == 0:
            worksheet.write(f"C{row}", "Насыщения не происходит", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas3_Int_0'])} < {str_cm(data['t_nas_treb_Int'])}", form_bold)
        row += 1

        worksheet.write_rich_string(f"B{row}", *get_rich_string("A(1-K|\u03b3|) - значение коэффициента при 3-ф КЗ в зоне действия защиты с учетом остаточной намагниченности", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['A3_Int']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Время до насыщения при 3-ф КЗ в зоне действия защиты c учетом остаточной намагниченности, мс", form_def)
        if data["t_nas3_Int"] > data["t_nas_treb_Int"]:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas3_Int'])} > {str_cm(data['t_nas_treb_Int'])}", form_def)
        elif data["t_nas3_Int"] == 0:
            worksheet.write(f"C{row}", "Насыщения не происходит", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas3_Int'])} < {str_cm(data['t_nas_treb_Int'])}", form_bold)
        row += 1

    #3-ф вне зоны действия защиты
    if data["ignore_Ext3"] == False:
        worksheet.write(f"B{row}", "A - значение коэффициента при 3-ф КЗ вне зоны действия защиты", form_def)
        worksheet.write(f"C{row}", str_cm(data['A3_Ext_0']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Время до насыщения при 3-ф КЗ вне зоны действия защиты, мс", form_def)
        if data["t_nas3_Ext_0"] > data["t_nas_treb_Ext"]:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas3_Ext_0'])} > {str_cm(data['t_nas_treb_Ext'])}", form_def)
        elif data["t_nas3_Ext_0"] == 0:
            worksheet.write(f"C{row}", "Насыщения не происходит", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas3_Ext_0'])} < {str_cm(data['t_nas_treb_Ext'])}", form_bold)
        row += 1

        worksheet.write_rich_string(f"B{row}", *get_rich_string("A(1-K|\u03b3|) - значение коэффициента при 3-ф КЗ вне зоны действия защиты с учетом остаточной намагниченности", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['A3_Ext']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Время до насыщения при 3-ф КЗ вне зоны действия защиты c учетом остаточной намагниченности, мс", form_def)
        if data["t_nas3_Ext"] > data["t_nas_treb_Ext"]:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas3_Ext'])} > {str_cm(data['t_nas_treb_Ext'])}", form_def)
        elif data["t_nas3_Ext"] == 0:
            worksheet.write(f"C{row}", "Насыщения не происходит", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas3_Ext'])} < {str_cm(data['t_nas_treb_Ext'])}", form_bold)
        row += 1

    #1-ф в зоне действия защиты
    if data["ignore_Int1"] == False:
        worksheet.write(f"B{row}", "A - значение коэффициента при 1-ф КЗ в зоне действия защиты", form_def)
        worksheet.write(f"C{row}", str_cm(data['A1_Int_0']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Время до насыщения при 1-ф КЗ в зоне действия защиты, мс", form_def)
        if data["t_nas1_Int_0"] > data["t_nas_treb_Int"]:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas1_Int_0'])} > {str_cm(data['t_nas_treb_Int'])}", form_def)
        elif data["t_nas1_Int_0"] == 0:
            worksheet.write(f"C{row}", "Насыщения не происходит", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas1_Int_0'])} < {str_cm(data['t_nas_treb_Int'])}", form_bold)
        row += 1

        worksheet.write_rich_string(f"B{row}", *get_rich_string("A(1-K|\u03b3|) - значение коэффициента при 1-ф КЗ в зоне действия защиты с учетом остаточной намагниченности", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['A1_Int']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Время до насыщения при 1-ф КЗ в зоне действия защиты c учетом остаточной намагниченности, мс", form_def)
        if data["t_nas1_Int"] > data["t_nas_treb_Int"]:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas1_Int'])} > {str_cm(data['t_nas_treb_Int'])}", form_def)
        elif data["t_nas1_Int"] == 0:
            worksheet.write(f"C{row}", "Насыщения не происходит", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas1_Int'])} < {str_cm(data['t_nas_treb_Int'])}", form_bold)
        row += 1

    #1-ф вне зоны действия защиты
    if data["ignore_Ext1"] == False:
        worksheet.write(f"B{row}", "A - значение коэффициента при 1-ф КЗ вне зоны действия защиты", form_def)
        worksheet.write(f"C{row}", str_cm(data['A1_Ext_0']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Время до насыщения при 1-ф КЗ вне зоны действия защиты, мс", form_def)
        if data["t_nas1_Ext_0"] > data["t_nas_treb_Ext"]:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas1_Ext_0'])} > {str_cm(data['t_nas_treb_Ext'])}", form_def)
        elif data["t_nas1_Ext_0"] == 0:
            worksheet.write(f"C{row}", "Насыщения не происходит", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas1_Ext_0'])} < {str_cm(data['t_nas_treb_Ext'])}", form_bold)
        row += 1

        worksheet.write_rich_string(f"B{row}", *get_rich_string("A(1-K|\u03b3|) - значение коэффициента при 1-ф КЗ вне зоны действия защиты с учетом остаточной намагниченности", *form_bundle))
        worksheet.write(f"C{row}", str_cm(data['A1_Ext']), form_def)
        row += 1

        worksheet.write(f"B{row}", "Время до насыщения при 1-ф КЗ вне зоны действия защиты c учетом остаточной намагниченности, мс", form_def)
        if data["t_nas1_Ext"] > data["t_nas_treb_Ext"]:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas1_Ext'])} > {str_cm(data['t_nas_treb_Ext'])}", form_def)
        elif data["t_nas1_Ext"] == 0:
            worksheet.write(f"C{row}", "Насыщения не происходит", form_def)
        else:
            worksheet.write(f"C{row}", f"{str_cm(data['t_nas1_Ext'])} < {str_cm(data['t_nas_treb_Ext'])}", form_bold)
        row += 1

    worksheet.merge_range(f"A{row_begin}:A{row-1}", "Проверка времени до насыщения", form_def_rotated)

    try:
        workbook.close()
    except xlsxwriter.exceptions.FileCreateError as error:
        messagebox.showerror(message="Close the Excel file first!")
        return
    
    startfile("log_short.xlsx") 
         
    
