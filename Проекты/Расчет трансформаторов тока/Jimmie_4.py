import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog as fd
import csv

# Some constants
K_STEPS = 16 # A number of steps for approximate calculation

class GUI(object):

    def __init__(self, master):

        # File reading
        self.frm_top = tk.Frame(master) # A frame for the top part
        self.read_filename = tk.StringVar() # A variable for string read file path
        self.read_filename.set("...")
        # File name of the current read file
        self.lbl_open_file_name = tk.Label(master=self.frm_top, height=3, textvariable=self.read_filename,
                                           anchor="w", wraplength=200, justify=tk.LEFT)
        self.lbl_open_file_name.grid(row=0, column=1, sticky="w")
        # A button for reading a file
        self.btn_read_from_file = tk.Button(
            master=self.frm_top,
            text="Открыть:",
            width=7, height=2,
            command=lambda: self.values_from_file()
        )
        self.btn_read_from_file.grid(row=0, column=0, padx=2, sticky="w")
        self.frm_top.grid(row=0, column=0, pady=0, columnspan=3, sticky="w")

        # Main parameters
        # A list of default main parameters, their names and units
        GUI_main_parameters = [
            ['Ктт', '600/5', 'о.е.'],
            ['R2тт', '0,35', 'Ом'],
            ['X2тт', '0', 'Ом'],
            ['cosφтт', '0,8', 'о.е.'],
            ['Lкаб', '105', 'м'],
            ['Sреле', '2', 'ВА'],
            ['cosφреле', '1', 'о.е.'],
            ['Kγ', '0,1', 'о.е.'],
        ]

        self.frm_left = tk.Frame(master) # A frame for the left side of GUI
        # A frame for main parameters
        self.frm_values = tk.Frame(master=self.frm_left, borderwidth=2, relief="groove")
        # Create all the labels
        for i, name in enumerate(GUI_main_parameters):
            lbl_param = tk.Label(master=self.frm_values, text=name[0])
            lbl_param.grid(row=i, column=0, sticky="e")

        # Create a dict for all the entries in the GUI
        self.entries = {}
        # Fill the entry list with the main params for now
        for i, name in enumerate(GUI_main_parameters):
            ent_value = tk.Entry(master=self.frm_values, width=7, borderwidth=2) # Make an entry object
            ent_value.grid(row=i, column=1, sticky='w') # Place it in the GUI
            ent_value.insert(0, name[1]) # Insert the default value
            self.entries.update({name[0]: ent_value}) # Update the dict with the short name and the object

        # Create all the unit labels
        for i, name in enumerate(GUI_main_parameters):
            lbl_unit = tk.Label(master=self.frm_values, text=name[2]) # Create a label
            lbl_unit.grid(row=i, column=2, sticky="w") # Place it

        self.frm_values.grid(row=0, column=0) # Place the whole unit stuff

        # 3ph fault data

        self.lbl_3ph = tk.Label(master=self.frm_left, text="3-ф КЗ")
        self.lbl_3ph.grid(row=1, column=0)
        # A frame for 3ph currents
        self.frm_currents_3ph = tk.Frame(master=self.frm_left, borderwidth=2, relief="groove")
        # Labels for GUI clarity
        self.lbl_Ikz3 = tk.Label(master=self.frm_currents_3ph, text="Iкз, А")
        self.lbl_Ikz3.grid(row=0, column=1)
        self.lbl_Tp3 = tk.Label(master=self.frm_currents_3ph, text="Tp, мс")
        self.lbl_Tp3.grid(row=0, column=2)
        # Labels for 3ph fault parameters
        self.lbl_vnutr_3ph = tk.Label(master=self.frm_currents_3ph, text="Внут.")
        self.lbl_vnutr_3ph.grid(row=1, column=0, sticky="e")
        self.lbl_vnesh_3ph = tk.Label(master=self.frm_currents_3ph, text="Внеш.")
        self.lbl_vnesh_3ph.grid(row=2, column=0, sticky="e")
        # A list of current parameters, their names and default values
        GUI_current_parameters_3ph = [
            ['Iкз3.внутр', '10278'],
            ['Iкз3.внеш', '977'],
            ['Tp3.внутр', '19'],
            ['Tp3.внеш', '21'],
        ]
        # We add all the current data entries into our self.entries dict and insert the default values
        for name in GUI_current_parameters_3ph:
            entry = tk.Entry(master=self.frm_currents_3ph, width=7, borderwidth=2) # Create an entry
            entry.insert(0, name[1]) # Insert the default value
            self.entries.update({name[0]: entry}) # Update the dict
        # Then we place all the enries into GUI
        self.entries['Iкз3.внутр'].grid(row=1, column=1, sticky="w")
        self.entries['Tp3.внутр'].grid(row=1, column=2, sticky="w")
        self.entries['Iкз3.внеш'].grid(row=2, column=1, sticky="w")
        self.entries['Tp3.внеш'].grid(row=2, column=2, sticky="w")

        self.frm_currents_3ph.grid(row=2, column=0) # Place the 3ph fault data

        # 1ph fault data

        self.lbl_1ph = tk.Label(master=self.frm_left, text="1-ф КЗ")
        self.lbl_1ph.grid(row=3, column=0)
        # A frame for 1ph fault data
        self.frm_currents_1ph = tk.Frame(master=self.frm_left, borderwidth=2, relief="groove")
        # Labels for GUI clarity
        self.lbl_Ikz1 = tk.Label(master=self.frm_currents_1ph, text="Iкз, А")
        self.lbl_Ikz1.grid(row=0, column=1)
        self.lbl_Tp1 = tk.Label(master=self.frm_currents_1ph, text="Tp, мс")
        self.lbl_Tp1.grid(row=0, column=2)
        # Labels for 1ph fault parameters
        self.lbl_vnutr_1ph = tk.Label(master=self.frm_currents_1ph, text="Внут.")
        self.lbl_vnutr_1ph.grid(row=1, column=0, sticky="e")
        self.lbl_vnesh_1ph = tk.Label(master=self.frm_currents_1ph, text="Внеш.")
        self.lbl_vnesh_1ph.grid(row=2, column=0, sticky="e")
        # A list of current parameters, their names and default values
        GUI_current_parameters_1ph = [
            ['Iкз1.внутр', '987'],
            ['Iкз1.внеш', '500'],
            ['Tp1.внутр', '27'],
            ['Tp1.внеш', '99'],
        ]
        # We add all the current data entries into our self.entries dict and insert the default values
        for name in GUI_current_parameters_1ph:
            entry = tk.Entry(master=self.frm_currents_1ph, width=7, borderwidth=2)  # Create an entry
            entry.insert(0, name[1])  # Insert the default value
            self.entries.update({name[0]: entry})  # Update the dict
        # Then we place all the enries into GUI
        self.entries['Iкз1.внутр'].grid(row=1, column=1, sticky="w")
        self.entries['Tp1.внутр'].grid(row=1, column=2, sticky="w")
        self.entries['Iкз1.внеш'].grid(row=2, column=1, sticky="w")
        self.entries['Tp1.внеш'].grid(row=2, column=2, sticky="w")

        self.frm_currents_1ph.grid(row=4, column=0)  # Place the 1ph fault data

        self.frm_left.grid(row=1, column=0, padx=4, pady=1, sticky="nw")   # Place the left frame

        # The right frame starts here
        self.frm_right = tk.Frame(master)
        # A frame for calculation parameters
        self.frm_val_boxes = tk.Frame(master=self.frm_right, pady=1)

        # Comboboxes for mode selection
        # A list of comboboxes, default parameters and options
        GUI_comboboxes = [
            ['Тип КЗ', '3-ф', ['3-ф', '1-ф']],
            ['Режим', 'Внутр.', ['Внутр.', 'Внеш.']],
            ['I(10%)', 'Внеш.', ['Внеш.', 'Внутр.', 'Iсз']],
            ['Схема ТТ', 'Y', ['Y', '△', 'неп.Y']]
        ]
        # Create a label for each box
        for i, name in enumerate(GUI_comboboxes):
            label = tk.Label(master=self.frm_val_boxes, text=name[0] + ':') # Create a label with the name
            label.grid(row=i, column=0, pady=1, sticky='w') # Place the label
        # Create comboboxes and add them to self.entries
        for i, name in enumerate(GUI_comboboxes):
            cbox = ttk.Combobox(master=self.frm_val_boxes, values=name[2], width=6) # Create a box
            cbox.grid(row=i, column=1, pady=1) # Place the box
            cbox.set(name[1]) # Set the default value
            self.entries.update({name[0]: cbox}) # Add the box to self.entries

        # Entries for calculation parameters
        # A list of parameter names and default values
        GUI_misc_parameters = [
            ['Kпер', '2'],
            ['Rпер(Ом)', '0,1'],
            ['Кол-во ТТ', '1'],
            ['Iсз (А)', ''],
            ['tнас (мс)', '30']
        ]
        # Create labels for the calculation parameters
        k = len(GUI_comboboxes) # A row offset
        for i, name in enumerate(GUI_misc_parameters):
            label = tk.Label(master=self.frm_val_boxes, text=name[0]) # Create a label
            label.grid(row=k+i, column=0, sticky='w') # Place the label
        # Create entries for the calculation parameters and add them to self.entries
        for i, name in enumerate(GUI_misc_parameters):
            entry = tk.Entry(master=self.frm_val_boxes, width=9, borderwidth=2) # Make an entry
            entry.grid(row=k+i, column=1, sticky='e') # Place it
            entry.insert(0, name[1]) # Insert the default value
            self.entries.update({name[0]: entry}) # Add it to self.entries
        self.entries['Iсз (А)'].configure(state='disabled') # Disable this entry by default
        self.entries['I(10%)'].bind('<<ComboboxSelected>>', self.update_Iras)

        self.frm_val_boxes.grid(row=0, column=0, pady=2, sticky="w") # Place the misc values frame
        # A button for approximate calculation
        self.btn_podgon = tk.Button(
            master=self.frm_right,
            text="Подгон!",
            width=14
            #command=lambda: [podgon()]
        )
        self.btn_podgon.grid(row=1, column=0, pady=2)

        # Calculation with fixed parameters
        # A frame for fixed calculation
        self.frm_fixed_podgon = tk.Frame(master=self.frm_right)
        # A list of fixed parameters and their default values
        GUI_fixed_parameters = [
            ['Sном', '0'],
            ['Kном', '0'],
            ['Sкаб', '0']
        ]
        # Add the fixed parameter labels
        for i, name in enumerate(GUI_fixed_parameters):
            label = tk.Label(master=self.frm_fixed_podgon, text=name[0]) # Add the label
            label.grid(row=i, column=0, sticky="w") # Place the label
        # Add the fixed parameter labels and add them to the self.entries
        for i, name in enumerate(GUI_fixed_parameters):
            entry = tk.Entry(master=self.frm_fixed_podgon, width=11, borderwidth=2) # Add an entry
            entry.grid(row=i, column=1) # Place the entry
            entry.insert(0, name[1]) # Insert the default value
            self.entries.update({name[0]: entry}) # Add the entry to self.entries

        # Manual Rcab enabling
        self.manual_Rcab = tk.IntVar()

        # A dict of checkbutton variables for later use
        self.checkbutton_vars = {
            'Ручное Rкаб': self.manual_Rcab
        }
        # Create a checkbox for enabling manual Rcab and add it tp self.entries
        chb_manual_Rcab = tk.Checkbutton(master=self.frm_fixed_podgon, text='Ручное Rкаб',
                                              variable=self.checkbutton_vars['Ручное Rкаб'],
                                              command=lambda: [
                                                  self.toggle_tk_object(
                                                      self.entries['Rкаб'],
                                                      self.entries['Sкаб'],
                                                      self.entries['Ручное Zнагр'])
                                              ],
                                              onvalue=1, offvalue=0)
        self.entries.update({'Ручное Rкаб': chb_manual_Rcab})
        self.entries['Ручное Rкаб'].grid(row=3, column=0, columnspan=2)
        # Create a label for manual Rcab entry
        self.lbl_manual_Rcab = tk.Label(master=self.frm_fixed_podgon, text='Rкаб')
        self.lbl_manual_Rcab.grid(row=4, column=0, sticky='w')
        # Create an entry for manual Rcab value and add it to self.entries
        ent_manual_Rcab = tk.Entry(master=self.frm_fixed_podgon, width=11, borderwidth=2)
        self.entries.update({'Rкаб': ent_manual_Rcab})
        self.entries['Rкаб'].grid(row=4, column=1)
        self.entries['Rкаб'].insert(0, '0')
        # Disable this entry by default
        self.entries['Rкаб'].configure(state='disabled')

        # Enabling of fixed Znagr
        # A variable for Znagr checkbutton
        self.manual_Znagr = tk.IntVar()
        self.checkbutton_vars.update({'Ручное Zнагр': self.manual_Znagr}) # Update the dict
        # Make a checkbutton
        chb_manual_Znagr = tk.Checkbutton(master=self.frm_fixed_podgon, text="Ручное Zнагр",
                                          variable=self.checkbutton_vars['Ручное Zнагр'],
                                          command=lambda: [
                                          self.toggle_tk_object(
                                              self.entries['Ручное Rкаб'],
                                              self.entries['Sкаб'],
                                              self.entries['Rнагр'],
                                              self.entries['Xнагр'])
                                          ],
                                          onvalue=1, offvalue=0)
        # Add the checkbutton to self.entries
        self.entries.update({'Ручное Zнагр': chb_manual_Znagr})
        self.entries['Ручное Zнагр'].grid(row=5, column=0, columnspan=2)
        # Make a label for Rнагр
        self.lbl_manual_Rnagr = tk.Label(master=self.frm_fixed_podgon, text='Rнагр')
        self.lbl_manual_Rnagr.grid(row=6, column=0, sticky='w')
        # Make an entry for Rnagr and add it to self.entries
        ent_manual_Rnagr = tk.Entry(master=self.frm_fixed_podgon, width=11, borderwidth=2)
        self.entries.update({'Rнагр': ent_manual_Rnagr})
        self.entries['Rнагр'].grid(row=6, column=1)
        self.entries['Rнагр'].insert(0, '0')
        self.entries['Rнагр'].configure(state='disabled') # Disable it by default
        # Make a label for Xнагр
        self.lbl_manual_Xnagr = tk.Label(master=self.frm_fixed_podgon, text='Xнагр')
        self.lbl_manual_Xnagr.grid(row=7, column=0, sticky='w')
        # Make an entry for Xнагр and add it to self.entries
        ent_manual_Xnagr = tk.Entry(master=self.frm_fixed_podgon, width=11, borderwidth=2)
        self.entries.update({'Xнагр': ent_manual_Xnagr})
        self.entries['Xнагр'].grid(row=7, column=1)
        self.entries['Xнагр'].insert(0, '0')
        self.entries['Xнагр'].configure(state='disabled') # Disable it by default

        self.frm_fixed_podgon.grid(row=2, column=0, pady=2, sticky="w")

        # Fixed calculation button
        self.btn_podgon = tk.Button(
            master=self.frm_left,
            text="Фиксированный",
            width=18)
        self.btn_podgon.grid(row=5, column=0, pady=2)

        # Button for plotting
        self.btn_plot = tk.Button(
            master=self.frm_left,
            text="График!",
            width=18)
        self.btn_plot.grid(row=6, column=0, pady=2)

        self.frm_right.grid(row=1, column=1, pady=2, sticky="nw")

        # A table for approximate calculation
        # Create an output frame
        self.frm_output = tk.Frame(master)
        GUI_output_header = ["k", "Sном", "Kном", "A", "tнас"]
        # Place all the header labels
        for i, header in enumerate(GUI_output_header):
            output_label = tk.Label(master=self.frm_output, text=header, borderwidth=2,
                                    width=7, height=1, relief="groove", anchor="n")
            output_label.grid(row=0, column=i)
        # Now we make a list of lists for output values
        self.approximate_output = []
        for i, header in enumerate(GUI_output_header):
            results = []
            # Create a frame to make it look legible
            row_frame = tk.Frame(master=self.frm_output, borderwidth=2, relief="groove")
            for k in range(K_STEPS):
                # Male an individual label
                output_label = tk.Label(master=row_frame, width=7, bd=0, text=str(i)+' '+str(k))
                output_label.grid(row=k+1, column=i) # Place the label in the table
                results.append(output_label) # Add it to the column
            row_frame.grid(row=1, column=i) # Place the whole row
            self.approximate_output.append(results) # Add the row to the list

        # Results for fixed saturation calculation
        # Make a frame for fixed output
        self.frm_output_fixed = tk.Frame(master=self.frm_output, borderwidth=2,
                                         relief="groove")
        # A list of fixed output label dict keys and their default text
        GUI_fixed_output = [
            ['Sном', 'Sном = ВА'],
            ['Kном', 'Kном = '],
            ['Sкаб', 'Sкаб = мм²'],
            ['Zнагр', 'Zнагр = Ом'],
            ['A', 'A = '],
            ['tнас', 'tнас = мс'],
            ['alpha', 'α'],
            ['blank', '']
        ]
        # A dictionary for all the labels
        self.fixed_output = {}
        # Now make all the labels and add them to self.fixed_output
        for i, label in enumerate(GUI_fixed_output):
            # Make a label
            output_label = tk.Label(master=self.frm_output_fixed, text=label[1],
                                    width=19, anchor='w')
            # To place the labels in two columns
            column = i % 2
            row = i // 2
            output_label.grid(row=row, column=column) # Place the label
            self.fixed_output.update({label[0]: output_label}) # Add it to the dict

        self.frm_output_fixed.grid(row=2, column=0, columnspan=5)

        # Results for fixed 10% basic calculations
        # A frame for 10% calculations
        self.frm_output_10 = tk.Frame(master=self.frm_output, borderwidth=2,
                                      relief="groove")
        # A list of 10% output label dict keys and their default text
        GUI_output_10 = [
            ['U2', 'U₂ = В'],
            ['Umax', 'Umax = '],
            ['Kтреб', 'Kтреб = '],
            ['Kфакт', 'Kфакт = ']
        ]
        # A dictionary for all the labels
        self.output_10 = {}
        # Now make all the labels and add them to self.output_10
        for i, label in enumerate(GUI_output_10):
            # Make a label
            output_label = tk.Label(master=self.frm_output_10, text=label[1],
                                    width=19, anchor='w')
            # To place the labels in two columns
            column = i % 2
            row = i // 2
            output_label.grid(row=row, column=column)  # Place the label
            self.output_10.update({label[0]: output_label})  # Add it to the dict

        self.frm_output_10.grid(row=3, column=0, columnspan=5)



        # Time settings
        # A frame for time settings
        self.frm_T_settings = tk.Frame(master=self.frm_output, borderwidth=2, relief="groove")
        # Add a checkbutton for enabling manual time settings
        self.manual_T = tk.IntVar()
        # Add the checkbutton to the dictionary
        self.checkbutton_vars.update({'Ручное задание времени': self.manual_T})
        chb_manual_T = tk.Checkbutton(master=self.frm_T_settings, text="Ручное задание времени",
                                      offvalue=0, onvalue=1, variable=self.checkbutton_vars['Ручное задание времени'],
                                      command=lambda: [
                                          self.toggle_tk_object(
                                              self.entries['Tнач'],
                                              self.entries['Tкон'],
                                              self.entries['Шаг'])
                                      ]
                                      )
        # Add the checkbutton to self.entries
        self.entries.update({'Ручное задание времени': chb_manual_T})
        self.entries['Ручное задание времени'].grid(row=0, column=0, columnspan=6, padx=0, sticky="w")
        # A list of default names and values
        GUI_T_settings = [
            ['Tнач', '0'],
            ['Tкон', '40'],
            ['Шаг', '0.1']
        ]
        # Now add the labels and corresponding entries
        for i, label in enumerate(GUI_T_settings):
            # Create the label
            time_label = tk.Label(master=self.frm_T_settings, width=5, text=label[0], anchor="e")
            time_label.grid(row=1, column=2*i)
            # Create an entry
            time_entry = tk.Entry(master=self.frm_T_settings, width=7, borderwidth=2)
            # Add it to the self.entries
            self.entries.update({label[0]: time_entry})
            self.entries[label[0]].grid(row=1, column=2*i+1)
            # Insert the default value
            self.entries[label[0]].insert(0, label[1])
            # All of them are disabled by default
            self.entries[label[0]].configure(state='disabled')

        self.frm_T_settings.grid(row=4, column=0, columnspan=5, pady=5, sticky="w")

        self.frm_output.grid(row=0, column=2, rowspan=4, pady=5, sticky="n")

        # Log settings
        self.frm_log = tk.Frame(master)
        # Make a dictionary for all the log checkbuttons. We don't add them to self.entries because they aren't saved
        # The full log checkbutton
        self.log_vars = {}
        self.print_log = tk.IntVar()
        self.log_vars.update({'Печать в отчет': self.print_log})
        # Create the checkbutton
        self.chb_print_log = tk.Checkbutton(master=self.frm_log, text="Печать в отчет",
                                            variable=self.log_vars['Печать в отчет'],
                                            command=lambda: [self.toggle_tk_object(self.chb_print_log_short)],
                                            onvalue=1, offvalue=0)
        self.chb_print_log.grid(row=0, column=0, padx=0, sticky="w")
        # The shortened log checkbutton
        self.print_log_short = tk.IntVar()
        self.log_vars.update({'Сокращенный отчет': self.print_log_short})
        self.chb_print_log_short = tk.Checkbutton(master=self.frm_log, text="Сокращенный отчет",
                                                  variable=self.log_vars['Сокращенный отчет'],
                                                  onvalue=1, offvalue=0)
        self.chb_print_log_short.grid(row=0, column=1, padx=0, sticky="w")
        self.chb_print_log_short.configure(state="disabled") # This one is disabled by default
        # The cursive checkbutton
        self.print_log_ital = tk.IntVar()
        self.log_vars.update({'Курсив': self.print_log_ital})
        self.chb_print_log_ital = tk.Checkbutton(master=self.frm_log, text="Курсив",
                                                 variable=self.log_vars['Курсив'],
                                                 onvalue=1, offvalue=0)
        self.chb_print_log_ital.grid(row=1, column=0, padx=0, sticky="w")

        self.frm_log.grid(row=3, column=0, columnspan=2, sticky="nw")

        # Saving button
        self.save_filename = tk.StringVar() # A variable for string save file path
        self.save_filename.set("...")
        self.frm_bottom = tk.Frame(master)
        self.lbl_save_file_name = tk.Label(master=self.frm_bottom, height=3, textvariable=self.save_filename,
                                           anchor="w", wraplength=200, justify=tk.LEFT)
        self.lbl_save_file_name.grid(row=0, column=1, sticky="w")
        self.btn_save_to_file = tk.Button(
            master=self.frm_bottom,
            text="Сохр.:",
            command=lambda: self.save_values(),
            width=7, height=2
        )
        self.btn_save_to_file.grid(row=0, column=0, padx=2, sticky="w")
        self.frm_bottom.grid(row=2, column=0, padx=0, pady=0, columnspan=3, sticky="nw")

    def values_from_file(self):
        """A function for reading saved values from a .csv file"""
        filetypes = [("CSV files", ".csv")] # Supported file types
        filename = fd.askopenfilename(
            title="Открыть файл исходных данных",
            filetypes=filetypes) # Get a path to the file
        if filename: # If the user didn't cancel
            self.read_filename.set(self.shorten_filename(filename)) # Update the GUI with a short path

            with open(filename, 'r', newline='\n', encoding="utf-8") as csv_open: # Open the .csv file
                reader = csv.reader(csv_open) # Read it
                for row in reader:
                    value = str(row[0]).split(' = ') # Split the row into name and values
                    GUI_element = self.entries[value[0]] # A GIU element which needs to change
                    GUI_element.configure(state='normal') # Enable all the elements back
                    if type(GUI_element) == tk.Entry: # Check if it's an entry
                        GUI_element.delete(0, tk.END) # Clear the entry
                        GUI_element.insert(0, value[1]) # Insert the new value
                    elif type(GUI_element) == ttk.Combobox:  # Check if it's a combobox
                        GUI_element.set(value[1])  # Set the box value
                    elif type(GUI_element) == tk.Checkbutton:  # Check if it's a checkbutton
                        checkbutton_var = self.checkbutton_vars[value[0]] # Find the corresponding variable
                        checkbutton_var.set(int(value[1])) # Set the box value
                csv_open.close()

            # Here we disable unneeded entries based on set values
            # A condition for 'Iсз (A)' entry
            if self.entries['I(10%)'].get() != 'Iсз':
                # If there doesn't need to be a value
                self.entries['Iсз (А)'].delete(0, tk.END) # Clear the entry
                self.entries['Iсз (А)'].configure(state='disabled') # Disable the element
            # A condition for 'Rкаб' entry
            if self.checkbutton_vars['Ручное Rкаб'].get() == 0:
                # If manual Rcab is disabled
                self.entries['Rкаб'].delete(0, tk.END) # Clear the entry
                self.entries['Rкаб'].configure(state='disabled') # Disable the entry
            else:
                # If manual Rcab is enabled
                self.entries['Sкаб'].configure(state='disabled') # Disable the 'Sкаб' entry
                self.entries['Ручное Zнагр'].configure(state='disabled') # Disable the 'Zнагр' checkbutton
            if self.checkbutton_vars['Ручное Zнагр'].get() == 0:
                # If manual Znagr is enabled
                self.entries['Rнагр'].delete(0, tk.END) # Clear the entries
                self.entries['Xнагр'].delete(0, tk.END)
                self.entries['Rнагр'].configure(state='disabled') # Disable the entries
                self.entries['Xнагр'].configure(state='disabled')
            else:
                # If manual Znagr is enabled
                self.entries['Sкаб'].configure(state='disabled')  # Disable the 'Sкаб' entry
                self.entries['Ручное Rкаб'].configure(state='disabled')  # Disable the 'Rкаб' checkbutton
            # A condition for time settings entries
            if self.checkbutton_vars['Ручное задание времени'].get() == 0:
                # If manual time settings are disabled
                self.entries['Tнач'].configure(state='disabled') # Disable the entries
                self.entries['Tкон'].configure(state='disabled')
                self.entries['Шаг'].configure(state='disabled')

    def save_values(self):
        """A function for saving all the values and settings to a .csv file"""
        filetypes = [("CSV files", ".csv")] # A list of compatible formats
        # Open the window and get the file path
        filename = fd.asksaveasfilename(
            title='Сохранить исходные данные',
            filetypes=filetypes,
            defaultextension=".csv")
        if filename: # If the user didn't cancel
            self.save_filename.set(self.shorten_filename(filename)) # Update the GUI with a short path

            with open(filename, 'w', newline='\n', encoding="utf-8") as csv_open: # Open the .csv file
                writer = csv.writer(csv_open, delimiter='=') # Make a writer
                for key in self.entries.keys():
                    # If you can get a string with just using .get() on an object
                    if type(self.entries[key]) == tk.Entry or type(self.entries[key]) == ttk.Combobox:
                        writer.writerow([key + ' ', ' ' + self.entries[key].get()])
                    # For checkbuttons which values are held in a separate dictionary
                    elif type(self.entries[key]) == tk.Checkbutton:
                        writer.writerow([key + ' ', ' ' + str(self.checkbutton_vars[key].get())])

                csv_open.close()

    def shorten_filename(self, filename):
        """Shorten the file path to '.../folder/filename.csv'"""
        pos_slash1 = filename.rfind("/") # Find the last /
        pos_slash2 = filename.rfind("/", 0, pos_slash1) # Find the second to last /
        short_name = "..." + filename[pos_slash2:len(filename)] # Crop the path
        return short_name

    def toggle_tk_object(self, *tk_objects):
        """Toggles the state of a Tkinter object"""

        for object in tk_objects:
            if object.cget("state") == "normal":
                object.configure(state="disabled")
            else:
                object.configure(state="normal")

    def update_Iras(self, event):
        """Updates the 'Iсз (А)' entry, when I(10%) is selected through GUI"""
        I10 = self.entries['I(10%)'].get()

        if I10 == 'Iсз':
            self.entries['Iсз (А)'].configure(state='normal')
        else:
            self.entries['Iсз (А)'].configure(state='disabled')



root_widget = tk.Tk()
new_app = GUI(root_widget)
root_widget.wm_title("Jimmie is back... All hail Jimmie!")
root_widget.resizable(width=False, height=False)
root_widget.mainloop()