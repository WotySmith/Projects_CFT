from math import sqrt
import csv
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog as fd
import locale
import math as m
import Excel_writer_new

from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)
import matplotlib.pyplot as plt

BIAS_STEPS = 16
BIAS_START = 5
T_STEP = 0.1
dT_START = 20
dT_STOP = 10
K_NOM_PRED = 200

locale.setlocale(locale.LC_ALL, "ru_RU.UTF8")

def select_read_file():
    """Запускает диалоговаое окно выбора .csv файла. Возвращает строку с путем
    или None, если не получилось"""
    filetypes = [("CSV files", ".csv")]
    filename = fd.askopenfilename(
        title="Открыть файл исходных данных",
        filetypes=filetypes)
    if filename:
        return fr"{filename}"
    else:
        return

def select_save_file(mode="parameters"):
    """Запускает диалоговое окно для сохранения .csv файла. Возвращает строку с путем
    или None, если не получилось"""
    filetypes = [("CSV files", ".csv")]
    if mode == "parameters":
        title = "Сохранить исходные данные:"
    elif mode == "log":
        title = "Сохранить расчетные значения:"
    else:
        return
    filename = fd.asksaveasfilename(
        title=title,
        filetypes=filetypes,
        defaultextension=".csv")
    if filename:
        return fr"{filename}"
    else:
        return

def shorten_filename(filename):
    """Сокращает строку с путем до двух последних слэшей"""
    pos_slash1 = filename.rfind("/")
    pos_slash2 = filename.rfind("/", 0, pos_slash1)
    short_name = "..." + filename[pos_slash2:len(filename)]
    return short_name

def read_file(file_name):
    """Читает .csv файл с данными в форме 'Переменная = значение'.
     Убирает пробелы и меняет запятые на точки. Возвращает list со значениями в форме str"""
    File_name_open = file_name
    try:
        with open(File_name_open, 'r', newline='', encoding="utf-8") as csv_open:
            values = []
            csvreader = csv.reader(csv_open, delimiter="=")
            for row in csvreader:
                if len(row) > 1:
                    value = row[1].replace(" ", "")
                    value = value.replace(",", ".")
                    values.append(value)
            csv_open.close()
            # values_currents = values[-8:]
            # values_TT = values[:-8]
            return values
    except FileNotFoundError:
        messagebox.showerror(message="Файл не найден!")
        return

def values_from_file(file_label):
    """Принимает в себя объект типа tk.Label, в который записано название открытого файла.
    Берет глобальный list entries, который составялется в main
    Напрямую проставляет в эти объекты все значения из файла"""

    file_name = select_read_file()
    values_file = read_file(file_name)
    if not values_file:
        return
    else:
        values = values_file
        file_label["text"] = shorten_filename(file_name)
    for i in range(len(values)):
        enteries[i].delete(0, tk.END)
        enteries[i].insert(0, values[i])

def update_values():
    """Читает значения из объектов глобального list enteries.
    Переводит все str значения во float или int.
    Возвращает dict с названиями переменных и их значениями"""

    values_names = ["I1", "I2", "R2tt", "X2tt", "cosf_tt", "L_cab", "S_rele",
                    "cosf_rele", "k_gamma",
                    "Ikz3_Int", "Tp3_Int", "Ikz3_Ext", "Tp3_Ext",
                    "Ikz1_Int", "Tp1_Int", "Ikz1_Ext", "Tp1_Ext"]
    values_str = [entry.get() for entry in enteries]

    slash_pos = values_str[0].find('/')
    if slash_pos == -1:
        messagebox.showerror(message="Проверь Ктт!\nНеобходимо записывать в форме I1/I2")
        return

    values_str.insert(0, values_str[0][:slash_pos])
    values_str[1] = values_str[1][slash_pos + 1:]

    values_named = dict(zip(values_names, values_str))

    for value in values_named.items():
        values_named.update({value[0]: value[1].replace(",", ".")})

    for value in values_named.items():
        if "." in value[1]:
            try:
                converted = float(value[1])
            except ValueError:
                messagebox.showerror(message=f"Проверь значение {value[0]}!")
                return
        else:
            try:
                converted = int(value[1])
            except ValueError:
                messagebox.showerror(message=f"Проверь значение {value[0]}!")
                return
        values_named.update({value[0]: converted})

    return values_named

def values_save_file(file_label):
    """Принимает объект типа tk.Label, в который записывается имя сохраненного файла.
    Вызывает select_save_file() и по полученному пути записывает исходные данные"""

    filename = select_save_file()
    file_label["text"] = shorten_filename(filename)

    with open(filename, 'w+', newline='\n', encoding="utf-8") as csv_open:
        writer = csv.writer(csv_open, delimiter='=')

        writer.writerow(["Параметры ТТ"])
        for i in range(len(parameters)):
            writer.writerow([parameters[i] + " ", " " + enteries[i].get()])

        writer.writerow([])

        writer.writerow(["Токи КЗ"])
        currents = ["Iкз3.внутр", "Tp3.внутр", "Iкз3.внеш", "Tp3.внеш",
                    "Iкз1.внутр", "Tp1.внутр", "Iкз1.внеш", "Tp1.внеш"]
        for i in range(len(currents)):
            j = i - 8
            writer.writerow([currents[i] + " ", " " + enteries[j].get()])

        csv_open.close()

def check_values_int(str_value):
    """Принимает str значение и переводит его в int.
    Если не получается, то возвращает -1.
    Я хз, почему -1, надо будет как-нибудь поправить"""

    try:
        value = int(str_value)
        return value
    except ValueError:
        value = -1
        return value

def check_values_float(str_value):
    """Преобразует строку в float. Если ValueError, возвращает -1"""
    str_value = str_value.replace(",", ".")
    try:
        value = float(str_value)
        return value
    except ValueError:
        value = -1
        return value

def get_KZ_type():
    """Обращается к глобальному объекту cmb_KZ_type и возвращает 3 или 1.
    Если не получается, возвращает None и выдает окошко с ошибкой"""

    KZ_type_str = cmb_KZ_type.get()
    if KZ_type_str == "3-ф":
        return 3
    elif KZ_type_str == "1-ф":
        return 1
    else:
        messagebox.showerror(message="Проверь тип КЗ!")
        return

def get_CON_type():
    """Обращается к глобальному объекту cmb_CON_select и возвращает star, triangle или part Y.
    Если не получается, возвращает None и выдает окошко с ошибкой"""

    CON_type_str = cmb_CON_select.get()
    if CON_type_str == "Y":
        return "star"
    elif CON_type_str == "\u25b3":
        return "triangle"
    elif CON_type_str == "неп. Y":
        return "part Y"
    else:
        messagebox.showerror(message="Проверь схему ТТ!")
        return

def get_WM():
    """Обращается к глобальному объекту cmb_WM_select и возвращает Ext или Int.
    Если не получается, возвращает None и выдает окошко с ошибкой"""

    WM_rus = cmb_WM_select.get()
    if WM_rus == "Внеш.":
        return "Ext"
    elif WM_rus == "Внутр.":
        return "Int"
    else:
        messagebox.showerror(message="Проверь режим!")
        return

def get_t_nas():
    """Работает с глобальным объектом ent_t_nas. Если в поле нет слэша,
    возвращает tuple (tнас, tнас). Если есть слэш, то возвращается tuple
    (tнас_внутр, tнас_внеш)"""

    t_nas_str = ent_t_nas.get()
    slash_pos = t_nas_str.find('/')
    if slash_pos == -1:
        t_nas = check_values_float(t_nas_str)
        if t_nas == -1:
            messagebox.showerror(message="Проверь tнас!")
            return
        else:
            return t_nas, t_nas
    else:
        t_nas_Int = check_values_float(t_nas_str[0:slash_pos])
        t_nas_Ext = check_values_float(t_nas_str[slash_pos + 1:len(t_nas_str)])
        if t_nas_Int == -1 or t_nas_Ext == -1:
            messagebox.showerror(message="Проверь tнас!")
            return
        else:
            return t_nas_Int, t_nas_Ext

def get_Num_TT():
    """Работает с глобальным объектом ent_Num_TT. Возвращает int.
    Если не получилось, то возвращает None и включает окошко с ошибкой"""
    Num_TT_str = ent_Num_TT.get()
    Num_TT = check_values_int(Num_TT_str)
    if Num_TT != -1:
        return Num_TT
    else:
        messagebox.showerror(message="Проверь количество ТТ!")
        return

def get_K10():
    """Работает с глобальным объектом ent_K10. Возвращает float.
    Если не получилось, то возвращает None и включает окошко с ошибкой"""

    K10_str = ent_K10.get()
    K10 = check_values_float(K10_str)
    if K10 != -1:
        return K10
    else:
        messagebox.showerror(message="Проверь Kпер!")
        return

def get_Rp():
    """Работает с глобальным объектом ent_Rp. Возвращает float.
    Если не получилось, то возвращает None и включает окошко с ошибкой"""
    Rp_str = ent_Rp.get()
    Rp = check_values_float(Rp_str)
    if Rp != -1:
        return Rp
    else:
        messagebox.showerror(message="Проверь Rпер!")
        return

def get_I_ras (*args):
    """Работает с глобальными объектоми cmb_Iras_select и ent_I_ras. Возвращает Ext, Int или Isz.
    Если не получается, то возвращает None и включает окошко с ошибкой\n
    *args нужны как костыль для изменения статуса ent_I_ras в реальном времени (см. задание chb_Iras_select)"""

    I_ras_mode = cmb_Iras_select.get()
    if I_ras_mode == "Внеш.":
        ent_I_ras.configure(state="disabled")
        if not args:
            return "Ext"
    if I_ras_mode == "Внутр.":
        ent_I_ras.configure(state="disabled")
        if not args:
            return "Int"
    elif I_ras_mode == "Iсз":
        ent_I_ras.configure(state="normal")
        if not args:
            return "Isz"
    else:
        messagebox.showerror(message="Проверь I(10%)!")
        return None

def get_T_ras(t_nas_treb, full = False):
    """Работает с глобальными объектами: chb_manual_T, ent_T_stop, ent_T_start, ent_T_step.
    На вход принимает значение tнас_треб и режим работы. В режиме работы full=True расчет
    время начала принимается равным 0.\n
    Проверяет, включено ли ручное задание времени расчета. Возвращает три значения времени в tuple:
    t_кон, t_нач и t_шаг.
    Если не получается, то возвращает None и включает окошко с ошибкой
    Пока я не разобрался, какой нормально брать t_нач, оно тупо ставится равным нулю"""

    manual_T_get = manual_T.get()
    if manual_T_get == 0:
        t_stop = t_nas_treb + dT_STOP
        ent_T_stop.configure(state="normal")
        ent_T_stop.delete(0, tk.END)
        ent_T_stop.insert(0, str(t_stop))
        ent_T_stop.configure(state="disabled")

        # t_start = t_nas_treb - dT_START if full != True else 0
        # if t_start < 0:
        #     t_start = 0
        t_start = 0
        ent_T_start.configure(state="normal")
        ent_T_start.delete(0, tk.END)
        ent_T_start.insert(0, str(t_start))
        ent_T_start.configure(state="disabled")

        t_step = T_STEP
        ent_T_step.configure(state="normal")
        ent_T_step.delete(0, tk.END)
        ent_T_step.insert(0, str(t_step))
        ent_T_step.configure(state="disabled")

    else:
        t_stop = check_values_float(ent_T_stop.get())
        t_start = check_values_float(ent_T_start.get())
        t_step = check_values_float(ent_T_step.get())
        if -1 in (t_stop, t_start, t_step):
            messagebox.showerror(message="Проверь время расчета!")
            return

        if t_stop <= t_nas_treb:
            t_stop = t_nas_treb + dT_STOP
            ent_T_stop.delete(0, tk.END)
            ent_T_stop.insert(0, str(t_stop))

    return t_stop, t_start, t_step

def toggle_tk_object (*tk_objects):
    """Работает с объектами TkInter. Меняет их состояние с noramal на disabled и наоборот"""

    for object in tk_objects:
        if object.cget("state") == "normal":
            object.configure(state="disabled")
        else:
            object.configure(state="normal")

def K_pr(values, WM, KZ_type, alpha, full=False):
    """Строит график Кпр(t), вызывая для каждой точки функцию max_Kpr().
    Тут же определяется точность расчета. Возвращает tuple из двух list"""
    if values == None:
        return
    Kpr = []
    times = []
    alpha = alpha * m.pi / 180

    Tp = values[f"Tp{KZ_type}_{WM}"]
    if Tp == 0:
        messagebox.showerror(message="Tp не может равнятся 0мс!")
        return

    t_nas_treb_Int, t_nas_treb_Ext = get_t_nas()
    if WM == "Int":
        t_nas_treb = t_nas_treb_Int
    elif WM == "Ext":
        t_nas_treb = t_nas_treb_Ext
    else:
        return

    t_stop, t, t_step = get_T_ras(t_nas_treb, full)

    if t_step >= 1:
        error = 0
    else:
        t_step_str = str(t_step)
        error = len(t_step_str) - t_step_str.find(".") - 1

    while t < t_stop + t_step:
        Kpr_t = max_Kpr(t, Tp, alpha)
        Kpr.append(Kpr_t)
        times.append(round(t, error))
        t += t_step
    return Kpr, times

def max_Kpr(t_nas, Tp, alpha):
    """Возвращает максимальное значение Кпр для времени tnas.\n
    Работает за счет того, что гоняет угол альфа по кругу"""
    if t_nas == 0:

        return 0
    w = 2 * m.pi / 20
    nu = 0
    Kpr_last = (m.sin(alpha) * m.cos(nu) * m.exp(-t_nas / Tp) +
                m.cos(alpha) * m.cos(nu) * w * Tp * (1 - m.exp(-t_nas / Tp))) - m.sin(w * t_nas + alpha + nu) + m.cos(
        alpha) * m.sin(nu)
    nu += 0.017 / 2 * m.pi

    while nu <= m.pi:
        Kpr_nu = (m.sin(alpha) * m.cos(nu) * m.exp(-t_nas / Tp) +
                  m.cos(alpha) * m.cos(nu) * w * Tp * (1 - m.exp(-t_nas / Tp))) - m.sin(w * t_nas + alpha + nu) + m.cos(
            alpha) * m.sin(nu)
        if Kpr_nu < Kpr_last:
            return abs(Kpr_last)
        else:
            Kpr_last = Kpr_nu
        nu += 0.017 / 2 * m.pi
    return Kpr_last

def t_nas_A(A, K_pr, times):
    """Функция для удобства. Возвращает первое значение t для которого А больше,
    чем Кпр для предыдущего шага"""

    t_nas = 0
    for i in range(len(times) - 1):
        if A < K_pr[i + 1]:
            t_nas = times[i]
            break
    return t_nas

def main():

    def Z_nagr(Pop_ras, L_cab, R_rele, X_rele, KZ_type, mode=None):
        """Возвращает tuple (Rнагр, Xнагр) или float Rкаб в зависимости от режима:\n
        mode = cable - возвращает только сопротивление кабеля\n
        mode = K10_partY - для расчета 10% неполной звезды, возвращает только Rнагр\n
        Если выставлен режим 'Ручное Zнагр', то возвращает значения из полей. Сопротивление
        кабеля при этом равно 999"""

        if manual_Znagr.get() == 1:
            if mode == 'cable':
                return 999
            R_nagr_ras = check_values_float(ent_manual_Rnagr.get())
            X_nagr_ras = check_values_float(ent_manual_Xnagr.get())
            return R_nagr_ras, X_nagr_ras

        CON_type = get_CON_type()
        Rp = get_Rp()

        Num_TT = get_Num_TT()
        if Num_TT == None:
            return

        R_cab_ras = 1 / 57 * L_cab / Pop_ras

        if mode == "cable":
            return R_cab_ras

        if mode == "K10_partY":
            R_nagr_ras = 1 / Num_TT * (2 * R_cab_ras + 2 * R_rele + Rp)
            return R_nagr_ras

        if KZ_type == 3:
            if CON_type == "star":
                R_nagr_ras = 1 / Num_TT * (R_cab_ras + R_rele + Rp)
                X_nagr_ras = 1 / Num_TT * X_rele
                return R_nagr_ras, X_nagr_ras
            if CON_type == "triangle":
                R_nagr_ras = 1 / Num_TT * (3 * R_cab_ras + 3 * R_rele + Rp)
                X_nagr_ras = 1 / Num_TT * 3 * X_rele
                return R_nagr_ras, X_nagr_ras
            if CON_type == "part Y":
                R_nagr_ras = 1 / Num_TT * (m.sqrt(3) * R_cab_ras + 2 * R_rele + Rp)
                X_nagr_ras = 1 / Num_TT * 2 * X_rele
                return R_nagr_ras, X_nagr_ras

        elif KZ_type == 1:
            if CON_type == "star":
                R_nagr_ras = 1 / Num_TT * (2 * R_cab_ras + R_rele + Rp)
                X_nagr_ras = 1 / Num_TT * X_rele
                return R_nagr_ras, X_nagr_ras
            if CON_type == "triangle":
                R_nagr_ras = 1 / Num_TT * (2 * R_cab_ras + 2 * R_rele + Rp)
                X_nagr_ras = 1 / Num_TT * 2 * X_rele
                return R_nagr_ras, X_nagr_ras
            if CON_type == "part Y":
                messagebox.showerror(message="Расчет для неполной звезды при 1ф КЗ не предусмотрен!")
        else:
            return

    def podgon():
        def p_cycle(A_treb, bias):
            for i in range(10, K_NOM_PRED * 10 + 5, 1):
                K_ras = i / 10
                S_ras = bias * K_ras
                Z_ras = S_ras / v["I2"] ** 2
                z2_ras = sqrt(
                    (v["R2tt"] + Z_ras * v["cosf_tt"]) ** 2 + (v["X2tt"] + Z_ras * sqrt(1 - v["cosf_tt"] ** 2)) ** 2)
                A = v["I1"] * K_ras * z2_ras / (v[f"Ikz{KZ_type}_{WM}"] * z2) * (1 - v["k_gamma"])
                if A > A_treb:
                    t_nas = t_nas_A(A, Kpr, times)
                    result = [bias, round(S_ras, 2), K_ras, round(A, 2), t_nas]
                    return result
                if (K_ras == K_NOM_PRED):
                    t_nas = t_nas_A(A, Kpr, times)
                    result = [bias, round(S_ras, 2), K_ras, round(A, 2), t_nas]
                    return result

        v = update_values()
        if v == None:
            return

        WM = get_WM()
        if WM == None:
            return
        KZ_type = get_KZ_type()

        if (v[f"Ikz{KZ_type}_{WM}"] == 0) or (v[f"Tp{KZ_type}_{WM}"] == 0):
            messagebox.showerror(message="Нет значений для расчетного режима!")
            return

        S_pop = check_values_float(ent_Pop_fixed.get())
        if S_pop == 0:
            S_pop = 2.5

        R_rele = v["S_rele"] / v["I2"] ** 2 * v["cosf_rele"]
        X_rele = v["S_rele"] / v["I2"] ** 2 * sqrt(1 - v["cosf_rele"] ** 2)
        R_nagr_ras, X_nagr_ras = Z_nagr(S_pop, v["L_cab"], R_rele, X_rele, KZ_type)
        z2 = sqrt((v["R2tt"] + R_nagr_ras) ** 2 + (v["X2tt"] + X_nagr_ras) ** 2)
        alpha = round(m.acos((v["R2tt"] + R_nagr_ras) / z2) / m.pi * 180, 2)

        Kpr_times = K_pr(v, WM, KZ_type, alpha)
        if Kpr_times == None:
            return
        Kpr, times = Kpr_times

        t_start = check_values_float(ent_T_start.get())
        t_stop = check_values_float(ent_T_stop.get())

        t_nas_treb_Int, t_nas_treb_Ext = get_t_nas()
        if WM == "Int":
            t_nas_treb = t_nas_treb_Int
        elif WM == "Ext":
            t_nas_treb = t_nas_treb_Ext
        else:
            return

        A_treb = 0
        for i in range(len(times)):
            K = Kpr[i]
            if K >= A_treb:
                if times[i] > t_nas_treb:
                    break
                else:
                    A_treb = K

        print_log_get = print_log.get()
        if print_log_get == 1:
            try:
                log_file_name = select_save_file(mode="log")
            except FileNotFoundError:
                messagebox.showerror(message="Файл не найден!")
                return

            with open(log_file_name, "a", newline="", encoding="UTF-8") as log:
                writer = csv.writer(log, delimiter='\t', lineterminator='\n')
                writer.writerow("")
                writer.writerow(header)
        i = 0
        for bias_int in range(BIAS_START, BIAS_STEPS + BIAS_START, 1):
            bias = bias_int / 10
            pdg = p_cycle(A_treb, bias)

            if print_log_get == 1:
                # print (*pdg, sep='\t')

                with open(log_file_name, "a", newline="", encoding="UTF-8") as log:
                    writer = csv.writer(log, delimiter='\t', lineterminator='\n')
                    writer.writerow(pdg)
                log.close()

            output_k[i]["text"] = f"{pdg[0]}"
            output_S[i]["text"] = f"{pdg[1]}"
            output_K[i]["text"] = f"{pdg[2]}"
            if pdg[3] < A_treb:
                output_A[i]["fg"] = "red"
            else:
                output_A[i]["fg"] = "black"
            output_A[i]["text"] = f"{pdg[3]}"
            if pdg[4] == 0:
                output_t_nas[i]["fg"] = "green"
                output_t_nas[i]["text"] = f">{t_stop}"
            elif pdg[4] == t_start:
                output_t_nas[i]["fg"] = "blue"
                output_t_nas[i]["text"] = f"<{t_start}"
            elif pdg[4] < t_nas_treb:
                output_t_nas[i]["fg"] = "red"
                output_t_nas[i]["text"] = f"{pdg[4]}"
            else:
                output_t_nas[i]["fg"] = "black"
                output_t_nas[i]["text"] = f"{pdg[4]}"
            i += 1


    def fixed_podgon():
        v = update_values()
        if v == None:
            return

        WM = get_WM()
        if WM == None:
            return
        KZ_type = get_KZ_type()

        if (v[f"Ikz{KZ_type}_{WM}"] == 0) or (v[f"Tp{KZ_type}_{WM}"] == 0):
            messagebox.showerror(message="Нет значений для расчетного режима!")
            return

        S_fixed = ent_S_fixed.get()
        if S_fixed == "0":
            S_ras = 0
            ign_S = 0
        else:
            S_ras = check_values_float(S_fixed)
            ign_S = 1
            if S_ras == -1:
                messagebox.showerror(message="Проверь значения!")
                return

        K_fixed = ent_K_fixed.get()
        if K_fixed == "0":
            K_ras = 0
            ign_K = 0
        else:
            K_ras = check_values_float(K_fixed)
            ign_K = 1
            if K_ras == -1:
                messagebox.showerror(message="Проверь значения!")
                return

        R_rele = v["S_rele"] / v["I2"] ** 2 * v["cosf_rele"]
        X_rele = v["S_rele"] / v["I2"] ** 2 * sqrt(1 - v["cosf_rele"] ** 2)

        if manual_Rcab.get() == 0:
            Pop_fixed = ent_Pop_fixed.get()
            if Pop_fixed == "0":
                Pop_ras = 2.5
                ign_Pop = 0
            else:
                Pop_ras = check_values_float(Pop_fixed)
                ign_Pop = 1
                if Pop_ras == -1:
                    messagebox.showerror(message="Проверь значения!")
                    return
        else:
            ign_Pop = 1
            Rcab_temp = check_values_float(ent_manual_Rcab.get())
            if Rcab_temp == -1:
                messagebox.showerror(message="Проверь значения!")
            else:
                Pop_ras = 1/57 * v["L_cab"]/Rcab_temp

        if manual_Znagr.get() == 1:
            ign_Pop = 1


        R_nagr_ras, X_nagr_ras = Z_nagr(Pop_ras, v["L_cab"], R_rele, X_rele, KZ_type)

        Z_ras = S_ras / v["I2"] ** 2
        z2_ras = sqrt((v["R2tt"] + Z_ras * v["cosf_tt"]) ** 2 + (v["X2tt"] + Z_ras * sqrt(1 - v["cosf_tt"] ** 2)) ** 2)
        z2_fact = sqrt((v["R2tt"] + R_nagr_ras) ** 2 + (v["X2tt"] + X_nagr_ras) ** 2)
        A = v["I1"] * K_ras * z2_ras / (v[f"Ikz{KZ_type}_{WM}"] * z2_fact) * (1 - v["k_gamma"])

        alpha_ras = round(m.acos((v["R2tt"] + R_nagr_ras) / z2_fact) / m.pi * 180, 2)

        Kpr_times = K_pr(v, WM, KZ_type, alpha_ras)
        if Kpr_times == None:
            return
        Kpr, times = Kpr_times

        t_nas_treb_Int, t_nas_treb_Ext = get_t_nas()
        if WM == "Int":
            t_nas_treb = t_nas_treb_Int
        elif WM == "Ext":
            t_nas_treb = t_nas_treb_Ext
        else:
            return

        A_treb = 0
        for i in range(len(times)):
            K = Kpr[i]
            if K >= A_treb:
                if times[i] > t_nas_treb:
                    break
                else:
                    A_treb = K

        i = 0
        Pop_list = [2.5, 4, 6, 8, 10, 12, 16, 20]
        while A < A_treb:
            if ign_S != 1:
                if 0.05 * S_ras >= 0.5 or S_ras == 0:
                    S_ras = S_ras + 0.5
                else:
                    S_ras = 1.05 * S_ras
            if ign_K != 1:
                if 0.05 * K_ras >= 0.5 or K_ras == 0:
                    K_ras = K_ras + 0.5
                else:
                    K_ras = 1.05 * K_ras
                if K_ras >= K_NOM_PRED:
                    K_ras = K_NOM_PRED
            if ign_Pop != 1 and i < len(Pop_list):
                Pop_ras = Pop_list[i]

            R_nagr_ras, X_nagr_ras = Z_nagr(Pop_ras, v["L_cab"], R_rele, X_rele, KZ_type)
            Z_ras = S_ras / v["I2"] ** 2
            z2_ras = sqrt((v["R2tt"] + Z_ras * v["cosf_tt"]) ** 2 + (v["X2tt"] + Z_ras * sqrt(1 - v["cosf_tt"] ** 2)) ** 2)
            z2_fact = sqrt((v["R2tt"] + R_nagr_ras) ** 2 + (v["X2tt"] + X_nagr_ras) ** 2)

            A = v["I1"] * K_ras * z2_ras / (v[f"Ikz{KZ_type}_{WM}"] * z2_fact) * (1 - v["k_gamma"])

            if ign_Pop != 1:
                alpha_ras = round(m.acos((v["R2tt"] + R_nagr_ras) / z2_fact) / m.pi * 180, 2)
                Kpr, times = K_pr(v, WM, KZ_type, alpha_ras)

                A_treb = 0
                for j in range(len(times)):
                    K = Kpr[j]
                    if K >= A_treb:
                        if times[j] > t_nas_treb:
                            break
                        else:
                            A_treb = K

            i += 1
            if i > 200:
                # print ("\nI think I'm stuck")
                break

        t_step = check_values_float(ent_T_step.get())
        t_step_str = str(t_step)
        pos_dec = t_step_str.find(".")
        if pos_dec == -1:
            error = 2
        else:
            error = len(t_step_str) - pos_dec - 1
            if error < 2:
                error = 2

        Z_nagr_ras = sqrt(R_nagr_ras ** 2 + X_nagr_ras ** 2)
        t_nas = t_nas_A(A, Kpr, times)
        result = [round(S_ras, 2), round(K_ras, 2), round(Pop_ras, 2), round(Z_nagr_ras, 3), round(A, error), t_nas]

        lbl_output_fixed_S["text"] = f"Sном = {result[0]} ВА"
        lbl_output_fixed_K["text"] = f"Kном = {result[1]}"
        lbl_output_fixed_Pop["text"] = f"Sкаб = {result[2]} мм\u00B2"
        lbl_output_fixed_R["text"] = f"Zнагр = {result[3]} Ом"
        if result[4] < A_treb:
            lbl_output_fixed_A["fg"] = "red"
        else:
            lbl_output_fixed_A["fg"] = "black"
        lbl_output_fixed_A["text"] = f"A = {result[4]}"

        t_start = check_values_float(ent_T_start.get())
        t_stop = check_values_float(ent_T_stop.get())
        t_nas_treb = check_values_float(ent_t_nas.get())
        if t_nas == 0:
            lbl_output_fixed_t_nas["fg"] = "green"
            lbl_output_fixed_t_nas["text"] = f"tнас > {t_stop} мс"
        elif t_nas == t_start:
            lbl_output_fixed_t_nas["fg"] = "blue"
            lbl_output_fixed_t_nas["text"] = f"tнас < {t_start} мс"
        elif t_nas < t_nas_treb:
            lbl_output_fixed_t_nas["fg"] = "red"
            lbl_output_fixed_t_nas["text"] = f"tнас = {result[5]} мс"
        else:
            lbl_output_fixed_t_nas["fg"] = "black"
            lbl_output_fixed_t_nas["text"] = f"tнас = {result[5]} мс"

        if v[f"Ikz{KZ_type}_Int"] >= v[f"Ikz{KZ_type}_Ext"]:
            U2max = round(sqrt(2) * 2 * v[f"Ikz{KZ_type}_Int"] * Z_nagr_ras / v["I1"] * v["I2"], 2)
            I_U2 = v[f"Ikz{KZ_type}_Int"]
        else:
            U2max = round(sqrt(2) * 2 * v[f"Ikz{KZ_type}_Ext"] * Z_nagr_ras / v["I1"] * v["I2"], 2)
            I_U2 = v[f"Ikz{KZ_type}_Ext"]

        WM_ras = get_I_ras()
        K10 = get_K10()
        calc10_2ph = False
        if WM_ras == "Ext":
            I_ras_10 = v[f"Ikz{KZ_type}_Ext"]
            K_fact = round(K_ras * (v["R2tt"] + Z_ras) / (v["R2tt"] + R_nagr_ras), 2)
            K_treb = round(K10 * I_ras_10 / v["I1"], 2)

        elif WM_ras == "Int":
            I_ras_10 = v[f"Ikz{KZ_type}_Int"]
            K_fact = round(K_ras * (v["R2tt"] + Z_ras) / (v["R2tt"] + R_nagr_ras), 2)
            K_treb = round(K10 * I_ras_10 / v["I1"], 2)

        elif WM_ras == "Isz":
            I_ras_10 = check_values_float(ent_I_ras.get()) * 1.1
            if I_ras_10 == -1:
                messagebox.showerror(message="Проверь Iрас!")
                return
            K_treb = round(K10 * I_ras_10 / v["I1"], 2)
            K_fact = round(K_ras * (v["R2tt"] + Z_ras) / (v["R2tt"] + R_nagr_ras), 2)

            CON_type = get_CON_type()
            if CON_type == "part Y":
                calc10_2ph = True
                R_nagr_ras_2 = Z_nagr(Pop_ras, v["L_cab"], R_rele, X_rele, KZ_type, mode="K10_partY")
                Z_nagr_ras_2 = sqrt(R_nagr_ras_2 ** 2 + X_nagr_ras ** 2)
                K_fact_2 = round(K_ras * (v["R2tt"] + Z_ras) / (v["R2tt"] + R_nagr_ras_2), 2)

        lbl_output_U2["text"] = f"U\u2082 = {U2max} В"
        if U2max < 1400:
            lbl_output_U2_cond["fg"] = "green"
            lbl_output_U2_cond["text"] = f"{U2max} < 1400"
        else:
            lbl_output_U2_cond["fg"] = "red"
            lbl_output_U2_cond["text"] = f"{U2max} > 1400"

        lbl_output_Ktreb["text"] = f"Kтреб = {K_treb}"
        if calc10_2ph == False:
            lbl_output_Kfact["text"] = f"Кфакт = {K_fact}"
            if K_fact > K_treb:
                lbl_output_K_cond["fg"] = "green"
                lbl_output_K_cond["text"] = f"{K_fact} > {K_treb}"
            else:
                lbl_output_K_cond["fg"] = "red"
                lbl_output_K_cond["text"] = f"{K_fact} < {K_treb}"
        else:
            lbl_output_Kfact["text"] = f"Кфакт = {K_fact_2}"
            if K_fact_2 > K_treb:
                lbl_output_K_cond["fg"] = "green"
                lbl_output_K_cond["text"] = f"{K_fact_2} > {K_treb}"
            else:
                lbl_output_K_cond["fg"] = "red"
                lbl_output_K_cond["text"] = f"{K_fact_2} < {K_treb}"

        lbl_output_fixed_alpha["text"] = f"\u03b1 = {alpha_ras}\u00b0"

        print_log_get = print_log.get()

        if print_log_get == 1:

            R_cab = round(Z_nagr(Pop_ras, v["L_cab"], R_rele, X_rele, KZ_type, mode="cable"), 3)

            excel_data_nas = {"A3_Int_0": None, "A3_Int": None, "t_nas3_Int_0": None, "tnas3_Int": None,
                              "ignore_Int3": None,
                              "A3_Ext_0": None, "A3_Ext": None, "t_nas3_Ex_0": None, "tnas3_Ext": None, "ignore_Ext3": None,
                              "A1_Int_0": None, "A1_Int": None, "t_nas1_Int_0": None, "tnas1_Int": None,
                              "ignore_Int1": None,
                              "A1_Ext_0": None, "A1_Ext": None, "t_nas1_Ex_0": None, "tnas1_Ext": None, "ignore_Ext1": None}

            excel_data_phase = {}

            excel_data_nas.update({f"A{KZ_type}_{WM}_0": A / (1 - v["k_gamma"])})
            excel_data_nas.update({f"A{KZ_type}_{WM}": A})
            Kpr, times = K_pr(v, WM, KZ_type, alpha_ras, full=True)
            excel_data_nas.update({f"t_nas{KZ_type}_{WM}_0": t_nas_A(excel_data_nas[f"A{KZ_type}_{WM}_0"], Kpr, times)})
            excel_data_nas.update({f"t_nas{KZ_type}_{WM}": t_nas_A(excel_data_nas[f"A{KZ_type}_{WM}"], Kpr, times)})
            excel_data_nas.update({f"ignore_{WM}{KZ_type}": False})
            plot_fixed(mode="stealth", plot_data=[
                Kpr, times,
                excel_data_nas[f"t_nas{KZ_type}_{WM}_0"],
                excel_data_nas[f"A{KZ_type}_{WM}_0"],
                WM, KZ_type, 0])
            plot_fixed(mode="stealth", plot_data=[
                Kpr, times,
                excel_data_nas[f"t_nas{KZ_type}_{WM}"],
                excel_data_nas[f"A{KZ_type}_{WM}"],
                WM, KZ_type, v["k_gamma"]])

            excel_data_phase.update({f"R{KZ_type}": R_nagr_ras})
            excel_data_phase.update({f"X{KZ_type}": X_nagr_ras})
            excel_data_phase.update({f"Z_nagr{KZ_type}": Z_nagr_ras})
            excel_data_phase.update({f"U2max{KZ_type}": U2max})
            excel_data_phase.update({f"I_U2_{KZ_type}": I_U2})
            excel_data_phase.update({f"I_ras_10_{KZ_type}": I_ras_10})
            excel_data_phase.update({f"Kfact{KZ_type}": K_fact})
            excel_data_phase.update({f"Ktreb{KZ_type}": K_treb})
            excel_data_phase.update({f"z2_fact{KZ_type}": z2_fact})
            excel_data_phase.update({f"ignore_K10_{KZ_type}": False})
            excel_data_phase.update({f"calc10_2ph": calc10_2ph})
            if calc10_2ph == True:
                excel_data_phase.update({"R2": R_nagr_ras_2})
                excel_data_phase.update({"Z_nagr2": Z_nagr_ras_2})
                excel_data_phase.update({"Kfact2": K_fact_2})

            for KZ_emp in [3, 1]:
                for WM_emp in ["Int", "Ext"]:
                    if excel_data_nas[f"A{KZ_emp}_{WM_emp}_0"] == None:
                        if v[f"Ikz{KZ_emp}_{WM_emp}"] != 0:
                            excel_data_nas.update({f"ignore_{WM_emp}{KZ_emp}": False})

                            R, X = Z_nagr(Pop_ras, v["L_cab"], R_rele, X_rele, KZ_emp)
                            z2 = sqrt((v["R2tt"] + R) ** 2 + (v["X2tt"] + X) ** 2)
                            A0 = v["I1"] * K_ras * z2_ras / (v[f"Ikz{KZ_emp}_{WM_emp}"] * z2)
                            A = A0 * (1 - v["k_gamma"])
                            alpha = round(m.acos((v["R2tt"] + R) / z2) / m.pi * 180, 2)

                            excel_data_nas.update({f"A{KZ_emp}_{WM_emp}_0": A0})
                            excel_data_nas.update({f"A{KZ_emp}_{WM_emp}": A})
                            Kpr, times = K_pr(v, WM_emp, KZ_emp, alpha, full=True)
                            excel_data_nas.update(
                                {f"t_nas{KZ_emp}_{WM_emp}_0": t_nas_A(excel_data_nas[f"A{KZ_emp}_{WM_emp}_0"],
                                                                      Kpr, times)})
                            excel_data_nas.update({f"t_nas{KZ_emp}_{WM_emp}":
                                                       t_nas_A(excel_data_nas[f"A{KZ_emp}_{WM_emp}"],
                                                               Kpr, times)})
                            plot_fixed(mode="stealth", plot_data=[
                                Kpr, times,
                                excel_data_nas[f"t_nas{KZ_emp}_{WM_emp}_0"],
                                excel_data_nas[f"A{KZ_emp}_{WM_emp}_0"],
                                WM_emp, KZ_emp, 0])
                            plot_fixed(mode="stealth", plot_data=[
                                Kpr, times,
                                excel_data_nas[f"t_nas{KZ_emp}_{WM_emp}"],
                                excel_data_nas[f"A{KZ_emp}_{WM_emp}"],
                                WM_emp, KZ_emp, v["k_gamma"]])
                        else:
                            excel_data_nas.update({f"ignore_{WM_emp}{KZ_emp}": True})

            op_KZ = 1 if KZ_type == 3 else 3
            if (excel_data_nas[f"ignore_Int{op_KZ}"] == False) or (excel_data_nas[f"ignore_Ext{op_KZ}"] == False):
                R_op, X_op = Z_nagr(Pop_ras, v["L_cab"], R_rele, X_rele, op_KZ)
                Z_nagr_op = sqrt(R_op ** 2 + X_op ** 2)
                z2 = sqrt((v["R2tt"] + R_op) ** 2 + (v["X2tt"] + X_op) ** 2)
                I_U2_op = max([v[f"Ikz{op_KZ}_Int"], v[f"Ikz{op_KZ}_Ext"]])
                U2max_op = round(sqrt(2) * 2 * I_U2_op * Z_nagr_op / v["I1"] * v["I2"], 2)

                excel_data_phase.update({f"R{op_KZ}": R_op})
                excel_data_phase.update({f"X{op_KZ}": X_op})
                excel_data_phase.update({f"Z_nagr{op_KZ}": Z_nagr_op})
                excel_data_phase.update({f"U2max{op_KZ}": U2max_op})
                excel_data_phase.update({f"I_U2_{op_KZ}": I_U2_op})
                excel_data_phase.update({f"z2_fact{op_KZ}": z2})

                if WM_ras == "Isz":
                    I_ras_10_op = I_ras_10
                    K_fact_op = round(K_ras * (v["R2tt"] + Z_ras) / (v["R2tt"] + R_op), 2)
                    K_treb_op = round(K10 * I_ras_10_op / v["I1"], 2)

                    excel_data_phase.update({f"ignore_K10_{op_KZ}": False})
                    excel_data_phase.update({f"I_ras_10_{op_KZ}": I_ras_10_op})
                    excel_data_phase.update({f"Kfact{op_KZ}": K_fact_op})
                    excel_data_phase.update({f"Ktreb{op_KZ}": K_treb_op})

                elif excel_data_nas[f"ignore_Ext{op_KZ}"] == False:
                    if WM_ras == "Ext":
                        I_ras_10_op = v[f"Ikz{op_KZ}_Ext"]
                        K_fact_op = round(K_ras * (v["R2tt"] + Z_ras) / (v["R2tt"] + R_op), 2)
                        K_treb_op = round(K10 * I_ras_10_op / v["I1"], 2)
                    elif WM_ras == "Int":
                        I_ras_10_op = v[f"Ikz{op_KZ}_Int"]
                        K_fact_op = round(K_ras * (v["R2tt"] + Z_ras) / (v["R2tt"] + R_op), 2)
                        K_treb_op = round(K10 * I_ras_10_op / v["I1"], 2)

                    excel_data_phase.update({f"ignore_K10_{op_KZ}": False})
                    excel_data_phase.update({f"I_ras_10_{op_KZ}": I_ras_10_op})
                    excel_data_phase.update({f"Kfact{op_KZ}": K_fact_op})
                    excel_data_phase.update({f"Ktreb{op_KZ}": K_treb_op})

                else:
                    excel_data_phase.update({f"ignore_K10_{op_KZ}": True})

            else:
                excel_data_phase.update({f"ignore_K10_{op_KZ}": True})

            sinf_tt = round(sqrt(1 - v["cosf_tt"] ** 2), 2)
            sinf_rele = round(sqrt(1 - v["cosf_rele"] ** 2), 2)

            CON_type = get_CON_type()
            Rp = get_Rp()

            Num_TT = get_Num_TT()
            if Num_TT == 1:
                Num_TT_mult = ""
            elif Num_TT == 2:
                Num_TT_mult = "0.5*()"
            else:
                messagebox.showerror(message="В отчете не больше двух ТТ!")
                return

            excel_data = {"S_ras": S_ras, "K_ras": K_ras, "Pop_ras": Pop_ras, "Num_TT_mult": Num_TT_mult,
                          "R_cab": R_cab, "R_rele": R_rele, "X_rele": round(X_rele, 2), "sinf_rele": sinf_tt,
                          "CON_type": CON_type, "Z_ras": Z_ras, "WM_ras": WM_ras, "z2_ras": round(z2_ras, 2),
                          "sinf_tt": sinf_tt, "K10": round(K10, 2), "Rp": round(Rp, 2),
                          "t_nas_treb_Int": t_nas_treb_Int, "t_nas_treb_Ext": t_nas_treb_Ext}

            for items in v.items():
                excel_data.update({items[0]: items[1]})
            for items in excel_data_nas.items():
                excel_data.update({items[0]: items[1]})
            for items in excel_data_phase.items():
                excel_data.update({items[0]: items[1]})

            print_log_short_get = print_log_short.get()
            ital = print_log_ital.get()
            if print_log_short_get == 1:
                Excel_writer_new.save_report_short(excel_data, ital)
            else:
                Excel_writer_new.save_report(excel_data, ital)

    def plot_fixed(mode=None, plot_data=None):
        if mode != "stealth":
            WM = get_WM()
            if WM == None:
                return
            KZ_type = get_KZ_type()

            t_nas_str = lbl_output_fixed_t_nas["text"]
            A_str = lbl_output_fixed_A["text"]
            alpha_str = lbl_output_fixed_alpha["text"]
            if t_nas_str[7:-3] == "" or A_str[3::] == "" or alpha_str[3:-3] == "":
                messagebox.showerror(message="Сначала сделай фиксированный расчет!")
                return
            t_nas = float(t_nas_str[7:-3])
            A = float(A_str[3::])
            alpha = float(alpha_str[3:-3])

            Kpr_times = K_pr(update_values(), WM, KZ_type, alpha, full=True)
            if Kpr_times == None:
                return
            else:
                Kpr = Kpr_times[0]
                times = Kpr_times[1]

        if mode == "stealth":
            Kpr, times, t_nas, A, WM, KZ_type, k_gamma = plot_data
            t_nas = round(t_nas, 2)
            A = round(A, 2)

        fig = plt.Figure()
        ax = fig.add_subplot()
        line = ax.plot(times, Kpr)
        ax.set(xlabel="t, мс", ylabel="Kп.р., о.е.")
        ax.set_ylim(bottom=0)
        ax.set_xlim(left=0, right=times[-1])

        if t_nas > times[0] and t_nas < times[-1]:
            ax.vlines(x=t_nas, ymin=0, ymax=A, color="red",
                      linestyles="dashed", label=f"tнас. = {t_nas} мс")
            ax.hlines(y=A, xmin=0, xmax=t_nas, color="orange",
                      linestyles="dashed", label=f"A = {A} о.е.")
        else:
            ax.hlines(y=A, xmin=0, xmax=times[-1], color="orange",
                      linestyles="dashed", label=f"A = {A} о.е.")
            ax.set_ylim(top=1.15 * A)
        ax.grid()
        ax.legend()

        if mode != "stealth":
            window_plot = tk.Toplevel(master=window)
            window_plot.title("Plot")

            canvas = FigureCanvasTkAgg(fig, master=window_plot)
            canvas.draw()

            toolbar = NavigationToolbar2Tk(canvas, window_plot, pack_toolbar=False)
            toolbar.update()

            button_quit = tk.Button(master=window_plot, text="Quit", command=window_plot.destroy)

            button_quit.pack(side=tk.BOTTOM)
            toolbar.pack(side=tk.BOTTOM, fill=tk.X)
            canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

            window_plot.mainloop()

        if mode == "stealth":
            graph_file = "Graph_" + WM + str(KZ_type) + "_gamma_" + str(k_gamma) + ".png"
            fig.savefig(graph_file, bbox_inches="tight")
            return

    global parameters

    parameters = ["Kтт", "R2тт", "X2тт", "cos\u03c6тт",
                  "Lкаб", "Sреле", "cos\u03c6реле", "K\u03b3"]
    default_values = read_file("default.csv")
    units = ["о.е.", "Ом", "Ом", "о.е", "м", "ВА", "о.е.", "о.е"]

    window = tk.Tk()
    window.title("Jimmie is back... All hail Jimmie!")
    window.resizable(width=False, height=False)

    # Чтение файла

    frm_top = tk.Frame(master=window)
    lbl_open_file_name = tk.Label(master=frm_top, height=3, text="default.csv",
                                  anchor="w", wraplength=200, justify=tk.LEFT)
    lbl_open_file_name.grid(row=0, column=1, sticky="w")

    btn_read_from_file = tk.Button(
        master=frm_top,
        text="Открыть:",
        width=7, height=2,
        command=lambda: values_from_file(lbl_open_file_name)
    )
    btn_read_from_file.grid(row=0, column=0, padx=2, sticky="w")
    frm_top.grid(row=0, column=0, pady=0, columnspan=3, sticky="w")

    # Исходные данные

    frm_left = tk.Frame(master=window)
    frm_values = tk.Frame(master=frm_left, borderwidth=2, relief="groove")
    for i in range(len(parameters)):
        lbl_param = tk.Label(master=frm_values, text=parameters[i])
        lbl_param.grid(row=i, column=0, sticky="e")

    global enteries
    enteries = []
    for i in range(len(parameters)):
        ent_value = tk.Entry(master=frm_values, width=7, borderwidth=2)
        enteries.append(ent_value)
        enteries[i].grid(row=i, column=1, sticky="w")
        enteries[i].insert(0, default_values[i])

    for i in range(len(parameters)):
        lbl_unit = tk.Label(master=frm_values, text=units[i])
        lbl_unit.grid(row=i, column=2, sticky="w")
    frm_values.grid(row=0, column=0)

    # Исходные данные 3-ф КЗ

    lbl_3ph = tk.Label(master=frm_left, text="3-ф КЗ")
    lbl_3ph.grid(row=1, column=0)

    frm_currents_3ph = tk.Frame(master=frm_left, borderwidth=2, relief="groove")

    lbl_Ikz3 = tk.Label(master=frm_currents_3ph, text="Iкз, А")
    lbl_Ikz3.grid(row=0, column=1)
    lbl_Tp3 = tk.Label(master=frm_currents_3ph, text="Tp, мс")
    lbl_Tp3.grid(row=0, column=2)

    lbl_vnutr_3ph = tk.Label(master=frm_currents_3ph, text="Внут.")
    lbl_vnutr_3ph.grid(row=1, column=0, sticky="e")
    ent_Ikz3_vnutr = tk.Entry(master=frm_currents_3ph, width=7, borderwidth=2)
    enteries.append(ent_Ikz3_vnutr)
    ent_Ikz3_vnutr.grid(row=1, column=1, sticky="w")
    ent_Ikz3_vnutr.insert(0, default_values[-8])
    ent_Tp3_vnutr = tk.Entry(master=frm_currents_3ph, width=7, borderwidth=2)
    enteries.append(ent_Tp3_vnutr)
    ent_Tp3_vnutr.grid(row=1, column=2, sticky="w")
    ent_Tp3_vnutr.insert(0, default_values[-7])

    lbl_vnesh_3ph = tk.Label(master=frm_currents_3ph, text="Внеш.")
    lbl_vnesh_3ph.grid(row=2, column=0, sticky="e")
    ent_Ikz3_vnesh = tk.Entry(master=frm_currents_3ph, width=7, borderwidth=2)
    enteries.append(ent_Ikz3_vnesh)
    ent_Ikz3_vnesh.grid(row=2, column=1, sticky="w")
    ent_Ikz3_vnesh.insert(0, default_values[-6])
    ent_Tp3_vnesh = tk.Entry(master=frm_currents_3ph, width=7, borderwidth=2)
    enteries.append(ent_Tp3_vnesh)
    ent_Tp3_vnesh.grid(row=2, column=2, sticky="w")
    ent_Tp3_vnesh.insert(0, default_values[-5])

    frm_currents_3ph.grid(row=2, column=0)

    # Исходные данные 1-ф КЗ

    lbl_1ph = tk.Label(master=frm_left, text="1-ф КЗ")
    lbl_1ph.grid(row=3, column=0)

    frm_currents_1ph = tk.Frame(master=frm_left, borderwidth=2, relief="groove")

    lbl_Ikz1 = tk.Label(master=frm_currents_1ph, text="Iкз, А")
    lbl_Ikz1.grid(row=0, column=1)
    lbl_Tp1 = tk.Label(master=frm_currents_1ph, text="Tp, мс")
    lbl_Tp1.grid(row=0, column=2)

    lbl_vnutr_1ph = tk.Label(master=frm_currents_1ph, text="Внут.")
    lbl_vnutr_1ph.grid(row=1, column=0, sticky="e")
    ent_Ikz1_vnutr = tk.Entry(master=frm_currents_1ph, width=7, borderwidth=2)
    enteries.append(ent_Ikz1_vnutr)
    ent_Ikz1_vnutr.grid(row=1, column=1, sticky="w")
    ent_Ikz1_vnutr.insert(0, default_values[-4])
    ent_Tp1_vnutr = tk.Entry(master=frm_currents_1ph, width=7, borderwidth=2)
    enteries.append(ent_Tp1_vnutr)
    ent_Tp1_vnutr.grid(row=1, column=2, sticky="w")
    ent_Tp1_vnutr.insert(0, default_values[-3])

    lbl_vnesh_1ph = tk.Label(master=frm_currents_1ph, text="Внеш.")
    lbl_vnesh_1ph.grid(row=2, column=0, sticky="e")
    ent_Ikz1_vnesh = tk.Entry(master=frm_currents_1ph, width=7, borderwidth=2)
    enteries.append(ent_Ikz1_vnesh)
    ent_Ikz1_vnesh.grid(row=2, column=1, sticky="w")
    ent_Ikz1_vnesh.insert(0, default_values[-2])
    ent_Tp1_vnesh = tk.Entry(master=frm_currents_1ph, width=7, borderwidth=2)
    enteries.append(ent_Tp1_vnesh)
    ent_Tp1_vnesh.grid(row=2, column=2, sticky="w")
    ent_Tp1_vnesh.insert(0, default_values[-1])

    frm_currents_1ph.grid(row=4, column=0)

    # Подгон

    frm_left.grid(row=1, column=0, padx=4, pady=1, sticky="nw")

    frm_right = tk.Frame(master=window)

    # Выбор режимов
    global cmb_KZ_type, cmb_CON_select, cmb_WM_select, ent_t_nas, ent_Num_TT, ent_K10, ent_Rp, cmb_Iras_select, ent_I_ras

    frm_val_boxes = tk.Frame(master=frm_right, pady=1)
    lbl_KZ_type = tk.Label(master=frm_val_boxes, text="Тип КЗ:")
    lbl_KZ_type.grid(row=0, column=0, pady=1, sticky="w")
    KZ_type_list = ["3-ф", "1-ф"]
    cmb_KZ_type = ttk.Combobox(master=frm_val_boxes, values=KZ_type_list, width=6)
    cmb_KZ_type.grid(row=0, column=1, pady=1)
    cmb_KZ_type.set("3-ф")

    lbl_WM_select = tk.Label(master=frm_val_boxes, text="Режим: ")
    lbl_WM_select.grid(row=1, column=0, pady=1, sticky="w")
    WM_select_list = ["Внутр.", "Внеш."]
    cmb_WM_select = ttk.Combobox(master=frm_val_boxes, values=WM_select_list, width=6)
    cmb_WM_select.grid(row=1, column=1, pady=1)
    cmb_WM_select.set("Внутр.")

    lbl_Iras_select = tk.Label(master=frm_val_boxes, text="I(10%):")
    lbl_Iras_select.grid(row=2, column=0, pady=1, sticky="w")
    Iras_select_list = ["Внеш.", "Внутр.", "Iсз"]
    cmb_Iras_select = ttk.Combobox(master=frm_val_boxes, values=Iras_select_list, width=6)
    cmb_Iras_select.grid(row=2, column=1, pady=1)
    cmb_Iras_select.set("Внеш.")
    cmb_Iras_select.bind("<<ComboboxSelected>>", get_I_ras)

    lbl_CON_select = tk.Label(master=frm_val_boxes, text="Схема ТТ:")
    lbl_CON_select.grid(row=3, column=0, sticky="w")
    CON_select_list = ["Y", "\u25b3", "неп. Y"]
    cmb_CON_select = ttk.Combobox(master=frm_val_boxes, values=CON_select_list, width=6)
    cmb_CON_select.grid(row=3, column=1, pady=1)
    cmb_CON_select.set("Y")

    # Расчетные значения

    lbl_K10 = tk.Label(master=frm_val_boxes, text="Кпер")
    lbl_K10.grid(row=4, column=0, sticky="w")
    ent_K10 = tk.Entry(master=frm_val_boxes, width=9, borderwidth=2)
    ent_K10.grid(row=4, column=1, sticky="e")
    ent_K10.insert(0, "2")

    lbl_Rp = tk.Label(master=frm_val_boxes, text="Rпер(Ом)")
    lbl_Rp.grid(row=5, column=0, sticky="w")
    ent_Rp = tk.Entry(master=frm_val_boxes, width=9, borderwidth=2)
    ent_Rp.grid(row=5, column=1, sticky="e")
    ent_Rp.insert(0, "0,1")

    lbl_Num_TT = tk.Label(master=frm_val_boxes, text="Кол-во ТТ")
    lbl_Num_TT.grid(row=6, column=0, sticky="w")
    ent_Num_TT = tk.Entry(master=frm_val_boxes, width=9, borderwidth=2)
    ent_Num_TT.grid(row=6, column=1, sticky="e")
    ent_Num_TT.insert(0, "1")

    lbl_I_ras = tk.Label(master=frm_val_boxes, text="Icз (А)")
    lbl_I_ras.grid(row=7, column=0, sticky="w")
    ent_I_ras = tk.Entry(master=frm_val_boxes, width=9, borderwidth=2)
    ent_I_ras.grid(row=7, column=1, sticky="e")
    ent_I_ras.configure(state="disabled")

    lbl_t_nas = tk.Label(master=frm_val_boxes, text="tнас (мс)")
    lbl_t_nas.grid(row=8, column=0, sticky="w")
    ent_t_nas = tk.Entry(master=frm_val_boxes, width=9, borderwidth=2)
    ent_t_nas.grid(row=8, column=1, sticky="e")
    ent_t_nas.insert(0, "30")

    frm_val_boxes.grid(row=0, column=0, pady=2, sticky="w")

    btn_podgon = tk.Button(
        master=frm_right,
        text="Подгон!",
        width=14,
        command=lambda: [podgon()]
    )
    btn_podgon.grid(row=1, column=0, pady=2)

    # Фиксированный подгон

    frm_fixed_podgon = tk.Frame(master=frm_right)
    lbl_S_fixed = tk.Label(master=frm_fixed_podgon, text="Sном")
    lbl_S_fixed.grid(row=0, column=0, sticky="w")
    ent_S_fixed = tk.Entry(master=frm_fixed_podgon, width=11, borderwidth=2)
    ent_S_fixed.grid(row=0, column=1)
    ent_S_fixed.insert(0, "0")
    lbl_K_fixed = tk.Label(master=frm_fixed_podgon, text="Kном")
    lbl_K_fixed.grid(row=1, column=0, sticky="w")
    ent_K_fixed = tk.Entry(master=frm_fixed_podgon, width=11, borderwidth=2)
    ent_K_fixed.grid(row=1, column=1)
    ent_K_fixed.insert(0, "0")
    lbl_Pop_fixed = tk.Label(master=frm_fixed_podgon, text="Sкаб")
    lbl_Pop_fixed.grid(row=2, column=0, sticky="w")
    ent_Pop_fixed = tk.Entry(master=frm_fixed_podgon, width=11, borderwidth=2)
    ent_Pop_fixed.grid(row=2, column=1)
    ent_Pop_fixed.insert(0, "0")

    #Включение фиксированного Rкаб
    global manual_Rcab
    manual_Rcab = tk.IntVar()
    chb_manual_Rcab = tk.Checkbutton(master=frm_fixed_podgon, text="Ручное Rкаб",
                                     variable=manual_Rcab,
                                     command=lambda: [toggle_tk_object(ent_manual_Rcab, ent_Pop_fixed, chb_manual_Znagr)],
                                     onvalue=1, offvalue=0)
    chb_manual_Rcab.grid(row=3, column=0, columnspan=2)
    lbl_manual_Rcab = tk.Label(master=frm_fixed_podgon, text='Rкаб')
    lbl_manual_Rcab.grid(row=4, column=0, sticky='w')
    ent_manual_Rcab = tk.Entry(master=frm_fixed_podgon, width=11, borderwidth=2)
    ent_manual_Rcab.grid(row=4, column=1)
    ent_manual_Rcab.insert(0, '0')
    ent_manual_Rcab.configure(state='disabled')

    #Включение фиксированного Znagr
    manual_Znagr = tk.IntVar()
    chb_manual_Znagr = tk.Checkbutton(master=frm_fixed_podgon, text="Ручное Zнагр",
                                     variable=manual_Znagr,
                                     command=lambda: [toggle_tk_object(ent_manual_Rnagr, ent_manual_Xnagr, ent_Pop_fixed, chb_manual_Rcab)],
                                     onvalue=1, offvalue=0)
    chb_manual_Znagr.grid(row=5, column=0, columnspan=2)

    lbl_manual_Rnagr= tk.Label(master=frm_fixed_podgon, text='Rнагр')
    lbl_manual_Rnagr.grid(row=6, column=0, sticky='w')
    ent_manual_Rnagr = tk.Entry(master=frm_fixed_podgon, width=11, borderwidth=2)
    ent_manual_Rnagr.grid(row=6, column=1)
    ent_manual_Rnagr.insert(0, '0')
    ent_manual_Rnagr.configure(state='disabled')

    lbl_manual_Xnagr= tk.Label(master=frm_fixed_podgon, text='Xнагр')
    lbl_manual_Xnagr.grid(row=7, column=0, sticky='w')
    ent_manual_Xnagr = tk.Entry(master=frm_fixed_podgon, width=11, borderwidth=2)
    ent_manual_Xnagr.grid(row=7, column=1)
    ent_manual_Xnagr.insert(0, '0')
    ent_manual_Xnagr.configure(state='disabled')

    frm_fixed_podgon.grid(row=2, column=0, pady=2, sticky="w")

    btn_podgon = tk.Button(
        master=frm_left,
        text="Фиксированный",
        width=18,
        command=lambda: [fixed_podgon()]
    )
    btn_podgon.grid(row=5, column=0, pady=2)

    # График

    btn_plot = tk.Button(
        master=frm_left,
        text="График!",
        width=18,
        command=plot_fixed
    )
    btn_plot.grid(row=6, column=0, pady=2)

    frm_right.grid(row=1, column=1, pady=2, sticky="nw")

    # Таблица результатов

    frm_output = tk.Frame(master=window)
    header = ["k", "Sном", "Kном", "A", "tнас"]

    lbl_output_k_header = tk.Label(master=frm_output, text=header[0], borderwidth=2,
                                   width=7, height=1, relief="groove", anchor="n")
    lbl_output_k_header.grid(row=0, column=0)
    output_k = []
    frm_output_k = tk.Frame(master=frm_output, borderwidth=2, relief="groove")
    for i in range(BIAS_STEPS):
        lbl_output_k = tk.Label(master=frm_output_k, width=7, bd=0)
        output_k.append(lbl_output_k)
        output_k[i].grid(row=i + 1, column=0)
    frm_output_k.grid(row=1, column=0)

    lbl_output_S_header = tk.Label(master=frm_output, text=header[1], borderwidth=2,
                                   width=7, height=1, relief="groove", anchor="n")
    lbl_output_S_header.grid(row=0, column=1)
    output_S = []
    frm_output_S = tk.Frame(master=frm_output, borderwidth=2, relief="groove")
    for i in range(BIAS_STEPS):
        lbl_output_S = tk.Label(master=frm_output_S, width=7, bd=0)
        output_S.append(lbl_output_S)
        output_S[i].grid(row=i + 1, column=0)
    frm_output_S.grid(row=1, column=1)

    lbl_output_K_header = tk.Label(master=frm_output, text=header[2], borderwidth=2,
                                   width=7, height=1, relief="groove", anchor="n")
    lbl_output_K_header.grid(row=0, column=2)
    output_K = []
    frm_output_K = tk.Frame(master=frm_output, borderwidth=2, relief="groove")
    for i in range(BIAS_STEPS):
        lbl_output_K = tk.Label(master=frm_output_K, width=7, bd=0)
        output_K.append(lbl_output_K)
        output_K[i].grid(row=i + 1, column=0)
    frm_output_K.grid(row=1, column=2)

    lbl_output_A_header = tk.Label(master=frm_output, text=header[3], borderwidth=2,
                                   width=7, height=1, relief="groove", anchor="n")
    lbl_output_A_header.grid(row=0, column=3)
    output_A = []
    frm_output_A = tk.Frame(master=frm_output, borderwidth=2, relief="groove")
    for i in range(BIAS_STEPS):
        lbl_output_A = tk.Label(master=frm_output_A, width=7, bd=0)
        output_A.append(lbl_output_A)
        output_A[i].grid(row=i + 1, column=0)
    frm_output_A.grid(row=1, column=3)

    lbl_output_t_nas_header = tk.Label(master=frm_output, text=header[4], borderwidth=2,
                                       width=7, height=1, relief="groove", anchor="n")
    lbl_output_t_nas_header.grid(row=0, column=4)
    output_t_nas = []
    frm_output_t_nas = tk.Frame(master=frm_output, borderwidth=2, relief="groove")
    for i in range(BIAS_STEPS):
        lbl_output_t_nas = tk.Label(master=frm_output_t_nas, width=7, bd=0)
        output_t_nas.append(lbl_output_t_nas)
        output_t_nas[i].grid(row=i + 1, column=0)
    frm_output_t_nas.grid(row=1, column=4)

    # Результаты фиксированные

    frm_output_fixed = tk.Frame(master=frm_output, borderwidth=2,
                                relief="groove")
    lbl_output_fixed_S = tk.Label(master=frm_output_fixed, text="Sном = ВА",
                                  width=19, anchor="w")
    lbl_output_fixed_S.grid(row=0, column=0)
    lbl_output_fixed_K = tk.Label(master=frm_output_fixed, text="Kном = ",
                                  width=19, anchor="w")
    lbl_output_fixed_K.grid(row=0, column=1)
    lbl_output_fixed_Pop = tk.Label(master=frm_output_fixed, text="Sкаб = мм\u00B2",
                                    width=19, anchor="w")
    lbl_output_fixed_Pop.grid(row=1, column=0)
    lbl_output_fixed_R = tk.Label(master=frm_output_fixed, text="Zнагр = Ом",
                                  width=19, anchor="w")
    lbl_output_fixed_R.grid(row=1, column=1)
    lbl_output_fixed_A = tk.Label(master=frm_output_fixed, text="A = ",
                                  width=19, anchor="w")
    lbl_output_fixed_A.grid(row=2, column=0)
    lbl_output_fixed_t_nas = tk.Label(master=frm_output_fixed, text="tнас = мс",
                                      width=19, anchor="w")
    lbl_output_fixed_t_nas.grid(row=2, column=1)
    lbl_output_fixed_alpha = tk.Label(master=frm_output_fixed, text="\u03b1 = ",
                                      width=19, anchor="w")
    lbl_output_fixed_alpha.grid(row=3, column=0)
    lbl_output_fixed_BLANK = tk.Label(master=frm_output_fixed, text="",
                                      width=19, anchor="w")
    lbl_output_fixed_BLANK.grid(row=3, column=1)

    frm_output_fixed.grid(row=2, column=0, columnspan=5)

    # Результаты 10%

    frm_output_10 = tk.Frame(master=frm_output, borderwidth=2,
                             relief="groove")
    lbl_output_U2 = tk.Label(master=frm_output_10, text="U\u2082 = В",
                             width=11, anchor="w")
    lbl_output_U2.grid(row=0, column=0)
    lbl_output_U2max = tk.Label(master=frm_output_10, text="Umax = 1400 В",
                                width=12, anchor="w")
    lbl_output_U2max.grid(row=0, column=1)
    lbl_output_U2_cond = tk.Label(master=frm_output_10, text="",
                                  width=14, anchor="w")
    lbl_output_U2_cond.grid(row=0, column=2)
    lbl_output_Ktreb = tk.Label(master=frm_output_10, text="Kтреб = ",
                                width=11, anchor="w")
    lbl_output_Ktreb.grid(row=1, column=0)
    lbl_output_Kfact = tk.Label(master=frm_output_10, text="Кфакт = ",
                                width=12, anchor="w")
    lbl_output_Kfact.grid(row=1, column=1)
    lbl_output_K_cond = tk.Label(master=frm_output_10, text="",
                                 width=14, anchor="w")
    lbl_output_K_cond.grid(row=1, column=2)

    frm_output_10.grid(row=3, column=0, columnspan=5)

    # Настройка времени
    global manual_T, ent_T_start, ent_T_stop, ent_T_step

    frm_T_settings = tk.Frame(master=frm_output, borderwidth=2, relief="groove")

    lbl_T_start = tk.Label(master=frm_T_settings, width=5, text="Tнач", anchor="e")
    lbl_T_start.grid(row=1, column=0)
    ent_T_start = tk.Entry(master=frm_T_settings, width=7, borderwidth=2)
    ent_T_start.grid(row=1, column=1)
    T_start = float(ent_t_nas.get()) - dT_START
    if T_start < 0:
        T_start = "0"
    else:
        T_start = str(T_start)
    ent_T_start.insert(0, T_start)
    ent_T_start.configure(state="disabled")

    lbl_T_stop = tk.Label(master=frm_T_settings, width=6, text="Tкон", anchor="e")
    lbl_T_stop.grid(row=1, column=2)
    ent_T_stop = tk.Entry(master=frm_T_settings, width=7, borderwidth=2)
    ent_T_stop.grid(row=1, column=3)
    T_stop = str(float(ent_t_nas.get()) + dT_STOP)
    ent_T_stop.insert(0, T_stop)
    ent_T_stop.configure(state="disabled")

    lbl_T_step = tk.Label(master=frm_T_settings, width=6, text="Шаг", anchor="e")
    lbl_T_step.grid(row=1, column=4)
    ent_T_step = tk.Entry(master=frm_T_settings, width=7, borderwidth=2)
    ent_T_step.grid(row=1, column=5)
    ent_T_step.insert(0, str(T_STEP))
    ent_T_step.configure(state="disabled")

    manual_T = tk.IntVar()
    chb_manual_T = tk.Checkbutton(master=frm_T_settings, text="Ручное задание времени",
                                  offvalue=0, onvalue=1, variable=manual_T,
                                  command=lambda: [
                                      toggle_tk_object(ent_T_start, ent_T_stop, ent_T_step)])
    chb_manual_T.grid(row=0, column=0, columnspan=6, padx=0, sticky="w")

    frm_T_settings.grid(row=4, column=0, columnspan=5, pady=5, sticky="w")

    # Печать лога

    frm_log = tk.Frame(master=window)
    print_log = tk.IntVar()
    chb_print_log = tk.Checkbutton(master=frm_log, text="Печать в отчет",
                                   variable=print_log,
                                   command=lambda: [toggle_tk_object(chb_print_log_short)],
                                   onvalue=1, offvalue=0)
    chb_print_log.grid(row=0, column=0, padx=0, sticky="w")

    print_log_short = tk.IntVar()
    chb_print_log_short = tk.Checkbutton(master=frm_log, text="Сокращенный отчет",
                                         variable=print_log_short, onvalue=1, offvalue=0)
    chb_print_log_short.grid(row=0, column=1, padx=0, sticky="w")
    chb_print_log_short.configure(state="disabled")

    print_log_ital = tk.IntVar()
    chb_print_log_ital = tk.Checkbutton(master=frm_log, text="Курсив",
                                        variable=print_log_ital, onvalue=1, offvalue=0)
    chb_print_log_ital.grid(row=1, column=0, padx=0, sticky="w")

    frm_log.grid(row=3, column=0, columnspan=2, sticky="nw")

    frm_output.grid(row=0, column=2, rowspan=4, pady=5, sticky="n")

    # Сохранение в файл

    frm_bottom = tk.Frame(master=window)
    lbl_save_file_name = tk.Label(master=frm_bottom, height=3, text="save.csv",
                                  anchor="w", wraplength=200, justify=tk.LEFT)
    lbl_save_file_name.grid(row=0, column=1, sticky="w")

    btn_save_to_file = tk.Button(
        master=frm_bottom,
        text="Сохр.:",
        width=7, height=2,
        command=lambda: values_save_file(lbl_save_file_name)
    )
    btn_save_to_file.grid(row=0, column=0, padx=2, sticky="w")
    frm_bottom.grid(row=2, column=0, padx=0, pady=0, columnspan=3, sticky="nw")

    tk.mainloop()

if __name__ == '__main__':
    main()