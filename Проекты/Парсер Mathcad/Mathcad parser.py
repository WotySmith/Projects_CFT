# -*- coding: utf-8 -*-
import xml.etree.ElementTree as ElemTree  # Import the XML ElementTree module aliased as ET
import io  # Import used for more advanced file writing
import os  # Used for file name handling
import numpy  # Used for arrays, and more advanced math manipulation
import base64  # Used for reading base64 encoded pictures
from tkinter.filedialog import askopenfilename  # Tkinter libs or selecting files in OS file picker
from tkinter import Tk, Frame, Button, Label, StringVar, BooleanVar, IntVar, Entry, Checkbutton, Scale  # Used for the GUI
# Other files
from symbol_parser import symbol_parser  # Import function which formats special charecters
from Tex_to_Excel import Tex_to_Excel

class ParseGUI(object):
    """Class used for the GUI
    The object parameter is the root widget for the Tkinter GUI
    """

    def __init__(self, master):
        """Constructor method

        :param master: A "master" wigdet
        """
        self.mainframe = Frame(master)  # Create a Frame child widget
        self.frame_top = Frame(self.mainframe) # Frame for labels
        self.frame_mid = Frame(self.mainframe)  # Frame for checkboxes
        self.frame_bottom = Frame(self.mainframe) # Frame for buttons
        self.frame_scale = Frame(self.mainframe) # Frame for the scale
        self.mainframe.pack()  # Make the widget visible

        self.path = ''  # Define default path for the .xmcd file
        self.texfile_path = ''  # Define path for the .tex file

        self.name = Label(self.mainframe, text="Welcome to Jimmie's Wonderland")  # Create a static text label
        self.name.pack(side="top")  # Make the widget visible and define location

        self.filename = StringVar()  # Create a dynamic string variable
        self.filename.set("Current selected file: none")  # Set the string value
        self.filename_label = Label(self.frame_top, textvariable=self.filename)  # Create a label with the dynamic var
        self.filename_label.pack()
        self.text_updater = Entry(self.frame_top, textvariable=self.filename)  # Create a Entry widget for auto updates

        self.status = StringVar()  # Used for displaying the status of the file operation
        self.status.set("Status: Not parsed")
        self.status_label = Label(self.frame_top, textvariable=self.status)
        self.status_label.pack()
        self.text_updater2 = Entry(self.frame_top, textvariable=self.status)
        self.frame_top.pack()

        self.auto_substitute = BooleanVar(value=True)
        self.chb_auto_substitute = Checkbutton(self.frame_mid, text="Substitute variables", offvalue=False,
                                               onvalue=True, variable=self.auto_substitute)
        self.chb_auto_substitute.pack(side='left')

        self.hide_units = BooleanVar(value=True)
        self.chb_hide_units = Checkbutton(self.frame_mid, text="Hide units", offvalue=False,
                                               onvalue=True, variable=self.hide_units)
        self.chb_hide_units.pack(side='left')

        self.cleanup = BooleanVar(value=True)
        self.chb_cleanup = Checkbutton(self.frame_mid, text="Cleanup", offvalue=False,
                                               onvalue=True, variable=self.cleanup)
        self.chb_cleanup.pack(side='left')
        self.frame_mid.pack()

        self.parse_file = Button(self.frame_bottom, text="Parse and save!", command=self.parse_file)  # Button for parsing
        self.parse_file.pack(side="right")

        self.parse_file = Button(self.frame_bottom, text="Open Excel file", command=self.open_file)
        self.parse_file.pack(side="right")

        self.select_file = Button(self.frame_bottom, text="Select file", command=self.select_file)  # Runs a class method
        self.select_file.pack(side="right")
        self.frame_bottom.pack()

        self.scale_name = Label(self.frame_scale, text="Decimal places: ", anchor='s') # A label for a scale
        self.scale_name.pack(side="left")

        self.scale_var = IntVar(value=3) # A scale for selecting number of decimal places
        self.scale = Scale(self.frame_scale, orient="horizontal", length=150, from_=0, to=5, variable=self.scale_var)
        self.scale.pack(side="right")
        self.frame_scale.pack()

    def select_file(self):  # Method used for selecting a file
        self.path = askopenfilename()  # Display native os file dialog for choosing a file
        self.filename.set("Current selected file: " + os.path.basename(self.path))  # Change the dynamic variable
        self.status.set("Status: Not parsed")  # Set status

    def open_file(self):  # Method used for opening the parsed LaTeX file
        self.texfile_path = os.path.dirname(self.path) + '/ParsedLatexFile/' + \
                            os.path.splitext(os.path.basename(self.path))[0] + '.tex'
        #if self.status.get() == "Status: File tried parsed! Look under the folder \ParsedLatexFile !":
            #os.system("start " + "\"\" \"" + self.texfile_path + "\"")

        Tex_to_Excel(self.texfile_path, self.cleanup.get()).write_Excel() # Convert Tex to Excel and open the file

    def parse_file(self):  # Method for parsing the chosen file
        # Make sure a file is selected and it is a Mathcad file before trying to parse it
        if self.filename.get() != 'Current selected file: none' and os.path.splitext(self.path)[1] == '.xmcd':
            self.status.set("Status: Tring to parse... (most files takes a few seconds)")
            MathcadXMLParser(self.path, self.auto_substitute.get(), self.hide_units.get(), self.scale_var.get())
            # Call the MathcadXMLParser class with the path
            self.status.set("Status: File tried parsed! Look under the folder \ParsedLatexFile !")
        # Display a error message to the user
        else:
            self.status.set("Status: You need to select a .xmcd (Mathcad) file!")


class MathcadXMLParser(object):
    """Class used for reading Mathcad files and writing LaTeX files

    Parses the results from the file reading and saving it
    Takes in a object (the file path)
    """

    def __init__(self, filename, auto_substitute=False, hide_units=False, dec_places=3):
        """The constructor method

        :param filename: A path to the targeted file
        """

        self.target_file = filename  # Define the filename as a class variable
        self.filename = os.path.basename(self.target_file).replace('.xmcd', '')  # Get the file name, remove extension
        math_tree_ok = True  # As a starting point, the file is ok

        try:  # Try to parse the file into a ElementTree
            self.math_tree = ElemTree.parse(self.target_file).getroot()  # Our full XML as an ElementTree

        except ElemTree.ParseError:  # Run if XML structure is corrupted
            print("Corrupted Mathcad document! Could not parse.")
            math_tree_ok = False  # Make sure rest of file is not tried parsed

        self.ml = "{http://schemas.mathsoft.com/math30}"  # Variable for namespaces as URI, used in prefixes
        self.ws = "{http://schemas.mathsoft.com/worksheet30}"
        self.un = "{http://schemas.mathsoft.com/units10}"

        self.units = {
            "ampere": "А",
            "ohm": "Ом",
            "volt": "В",
            "meter": "м",
            "watt": "Вт",
            "kilovolt": "кВ",
            "second": "с",
            "millimeter": "мм",
            "degC": "°C",
            "degree": "°"
        }

        self.output_folder = os.path.dirname(self.target_file) + "/ParsedLatexFile"
        if not os.path.exists(self.output_folder):  # If this folder doesn't exist
            os.makedirs(self.output_folder)  # Create it

        # Open a new tex file for writing. Encoding for Danish chars etc. Save in a folder
        self.tex_file = io.open(self.output_folder + "/" + self.filename + '.tex', 'w', encoding="utf-8")

        # Standard LaTeX document info as strings
        self.start_latex_doc = ""

        self.end_latex_doc = ""

        self.matrix_array = []  # Array to use for multiple values in matrixes

        self.current_region_no = 0  # To keep track of regions in mathcad file, as they are being read

        self.debug = False  # Toggle debug messages
        self.hide_units = hide_units # Don't show units in intermediary calculations controlled by GUI

        self.auto_substitute = auto_substitute # Toggle substituting values of the variables controlled by GUI
        self.variables = {} # A dictionary for putting in names and values of the variables

        self.dec_places = dec_places # A number of decimal places controlled by GUI

        if math_tree_ok:  # Only run if file isn't corrupted
            self.main()  # Run main method

    def math_reader(self, elem):
        """Method for getting the math out of the XML file

        This method is used for reading the XML file (the ElementTree).
        Generally speaking this method gathers the data, the
        latex_formatter method uses to format the data into LaTeX.
        So no LaTeX-specific formatting is done in this method.
        Recursive method for efficiency and simplicity.

        This method is used for math Mathcad regions.
        :param elem: A ElementTree
        :return: Full LaTeX formatted math expressions
        """
        elem.tag = elem.tag.replace(self.ml, "")  # Only leave operator name left without prefix
        elem.tag = elem.tag.replace(self.un, "")
        #print(elem.tag)

        if elem.tag == "apply":  # If current Element's tag is apply
            # We need two cases; One for apply tag which includes a operator, and everything else
            # The first will always be the operator, if there's one
            if self.debug: print("Apply tag found")

            # Either there's a operator (2 or 3 childs, first element is always the operator)
            if bool(elem[0]) is False and elem[0].text is None:  # Checks if the elem has children and doesn't have text
                if self.debug:  # Only prints debug messages if debug = True
                    print("Apply tag includes a operator")

                if len(elem) == 3:  # Either there's 3 parts (normal mathemathical expression)
                    if self.debug:
                        print("len(elem)", len(elem))
                    val1 = self.math_reader(elem[1])  # Call this method again with 2nd child again to get first value
                    val2 = self.math_reader(elem[2])  # Call this method again with 3rd child again to get second value
                    # Return the formatted result (by calling math_formatter), to the original caller of this method
                    return self.latex_formatter(elem[0].tag, val1, val2)  # Sends the operator and the two values

                elif len(elem) == 2:  # Used for other operators where there's only "two" parts
                    if self.debug: print("len(elem)", len(elem))
                    val1 = self.math_reader(elem[1])  # Call this method again, skip the operator tag elem
                    return self.latex_formatter(elem[0].tag, val1)  # Get first child's tag which is the operator

            # ToDo: Make a more general way of handling apply tags?
            # Or there's no operator - this is the case for ex cos(x)
            # One of the 2 childs, must have children too, or text
            elif bool(elem[0]) or bool(elem[1]) or elem[0].text or elem[1].text is not None:
                val1 = self.math_reader(elem[0])
                val2 = self.math_reader(elem[1])
                return self.latex_formatter(elem.tag, val1, val2)

        # ToDo?: The following else tag checks, are only for non-operator tags, elements with no children or text
        elif elem.tag == "parens":  # Handle parenteses
            return self.latex_formatter(elem.tag, self.math_reader(elem[0]))  # Only 1 value between parenteses

        # Current tag is some kind of equal sign
        elif elem.tag == "define" or elem.tag == "symEval" or elem.tag == "globalDefine":
            if self.debug: print("A type of equal expression found.")
            if len(elem) < 3: #Regular equal sign
                parced_elem0 = self.math_reader(elem[0])
                parced_elem1 = self.math_reader(elem[1])
                if self.auto_substitute and elem.tag == 'define':
                    # If substitution is toggled and we're dealing with a definition
                    end_pos = parced_elem1.rfind('=') # Find the last =
                    if end_pos != -1:
                        # If found remove the '= ' and get the result
                        var_value = parced_elem1[end_pos:].strip('= ')
                    else:
                        # Else just get the number
                        var_value = parced_elem1
                    if self.hide_units: # If hide units is toggled
                        unit_pos = var_value.rfind('·') # Find where the unit starts
                        if unit_pos != -1: # And cut it off
                            self.variables.update({parced_elem0: var_value[:unit_pos]})
                        else:
                            self.variables.update({parced_elem0: var_value})
                    else:
                        self.variables.update({parced_elem0: var_value})
                        # Add the name and the value to the variables dict for later replacement
                return self.latex_formatter(elem.tag, parced_elem0, parced_elem1)
            else: #In case there's a command tag inside
                return self.latex_formatter(elem.tag, self.math_reader(elem[0]), self.math_reader(elem[2]))

        # Eval works like normal equal operator, if the result is not defined using "define"
        # ToDo: add a way to handle non-user defined units in the result, and possibly add library to format to LaTeX?
        elif elem.tag == "eval":
            # Either the evaluated expression doesn't have a user defined unit in the result
            if len(elem) == 2:
                return self.latex_formatter(elem.tag, self.math_reader(elem[0]), self.math_reader(elem[1]))

            # Or it does, and we can use that as the unit
            elif len(elem) == 3:
                value_and_unit = self.math_reader(elem[2]) + '·' + self.math_reader(elem[1])
                return self.latex_formatter(elem.tag, self.math_reader(elem[0]), value_and_unit)

        elif elem.tag == "provenance":  # Interesting Mathcad structure handled here
            return self.math_reader(elem[len(elem) - 1])  # Simply call this method again with the last child element

        elif elem.tag == "id":  # Current tag is pure text
            if self.debug: print("Text found:", elem.text)
            return self.latex_formatter(elem.tag, elem)  # Call external function with text

        elif elem.tag == "real":  # Current is a real number
            if self.debug: print("Number found:", elem.text)
            number = numpy.round(float(elem.text),self.dec_places)
            # A check to prevent rounding small numbers to 0
            if len(elem.text) > 1 and number == 0:
                string_number = elem.text
            else:
                string_number = str(number)
            if string_number.endswith(".0"):
                return string_number.replace(".0", "")  # Simply return the value as string
            else:
                return string_number.replace(".", ",")

        elif elem.tag == "imag": # Current is an imaginary number
            if self.debug: print("Imaginary number found")
            number = numpy.round(float(elem.text), 2)
            string_number = str(number)
            if string_number.endswith(".0"):
                return "j" + string_number.replace(".0", "")  # Simply return the value as string
            else:
                return "j" + string_number.replace(".", ",")

        # Result is used with equal signs, boundVars is used for special variables
        # degree is used for n'te degree derivatives
        # symResult is a result from a symbolic evaluation
        elif elem.tag == "result" or elem.tag == "boundVars" or elem.tag == "degree" or elem.tag == "symResult":
            if self.debug: print("Result found.")
            return self.math_reader(elem[0])
            # return elem[0][0].text  # Simply return the value as string

        elif elem.tag == "vectorize":  # Current tag is a vector notation
            if self.debug: print("Vector found.")
            return self.latex_formatter(elem.tag, self.math_reader(elem[0]))

        elif elem.tag == "matrix":  # Currrent tag is a matrix
            if self.debug: print("Matrix found.")

            # Run through every entity in the matrix, and add it to a list
            for entity in elem:
                # Convert every entity in to string, for supporting advanced expressions in matrix
                self.matrix_array.append(str(self.math_reader(entity)))  # Recursive, to handling everything :D

            numpy_matrix_array = numpy.array(self.matrix_array)  # Convert from list to a numpy array for manipulation
            array_dimensions = (int(elem.attrib["rows"]), int(elem.attrib["cols"]))  # Grap the ints from the attribute
            numpy_matrix_array = numpy.reshape(numpy_matrix_array, array_dimensions)  # Transform from flat array
            if array_dimensions[0] > 1 and array_dimensions[1] > 1:  # Only transpose if there's more than 1 col or row
                numpy_matrix_array = numpy.transpose(numpy_matrix_array)  # Flip array around to fit LaTeX structure

            self.matrix_array = []  # Reset matrix array
            return self.latex_formatter(elem.tag, numpy_matrix_array, array_dimensions)  # Send array and the dimensons

        elif elem.tag == "placeholder":  # Current tag is just a placeholder
            if self.debug:
                print("Empty placeholder found.")
            return " "  # Return space

        elif elem.tag == "lambda" or elem.tag == "bounds":
            # lambda is used for both derivative and integral (+ more?!).
            # bounds is used for integral with limits
            # Therefore the latex_formatter must handle it, we can't go backwards in elements?
            return elem

        elif elem.tag == "unitedValue" or elem.tag == "unitOverride":  # Value and/or unit
            if self.debug:
                print("Unit stuff found.")
            if len(elem) < 2:
                return self.math_reader(elem[0])
            else:
                return self.latex_formatter(elem.tag, self.math_reader(elem[0]), self.math_reader(elem[1][0]))

        if elem.tag == "unitReference": # If elem is a special unit reference
            if self.debug:
                print("Unit found.")
            if elem.attrib["unit"] in self.units.keys():
                return self.units[elem.attrib["unit"]]
            else:
                print("Unsupported unit found:", elem.attrib["unit"])
                return "UNKNOWN UNIT"

        if elem.tag == "complex": # If elem is a complex number
            if self.debug:
                print("Complex number found")
            return self.latex_formatter("complex", self.math_reader(elem[0]), self.math_reader(elem[1]))

        elif elem.tag == "function":
            if self.debug:
                print("Function found.")
            return self.latex_formatter(elem.tag, self.math_reader(elem[0]), self.math_reader(elem[1]))

        else:  # For unsupported tags
            print("Error, non-supported tag found at region",
                  self.current_region_no)  # Print the problematic region number
            print("Current Elem.tag:", elem.tag)  # Debug message

    def text_reader(self, elem):
        """Parses Mathcad text regions paragraph by paragraph.
        It loops through each paragraph, and add new
        :param elem: The ElementTree that contains text, structure: <text> ... </text>
        :return: Formatted text with escape chars
        """
        text = ""  # Variable to store found text

        if self.debug:
            print("Type: Text region.")

        # For each paragraph <p>...</p>, in the text "element" object:  <text> ... </text>
        for i, paragraph in enumerate(elem, start=1):

            # Add text from start of paragraph
            if paragraph.text is not None and ("\t" not in paragraph.text):
                text += paragraph.text

            # If paragraph has children (can be multiple lines of text, math region, or bold/italic etc text)
            text += self.text_piece_formatter(paragraph)

            # Add tailing text for paragraph
            if (paragraph.tail is not None) and ("\t" not in paragraph.tail):
                text += paragraph.tail

            # Add newline after every paragraph that isn't the last
            if i <= len(elem):
                text += "\n"

            # For the last paragraph (no newline)
            else:
                text += ""

        # Make sure symbols are parsed correctly
        return symbol_parser(text, False)

    def text_piece_formatter(self, paragraph_elem):
        """Method to handle pieces of text in a paragraph.
        Can be boldface, italic etc, with different combinations of these
        We use depth first search to find text and format it using latex syntax.

        :param paragraph_elem: The ElementTree that contains a line of text to format
        :return: String with formatted using LaTeX syntax"""

        # String to store our text
        text = ""

        # Mapping from Mathcad xml tags to latex
        emphasis_latex = {
            "b": "?",
            "i": "\\textit",
            "u": "\\underline",
            "sup": "\\textsuperscript"
        }

        # Loop through each piece of text in the paragraph. The order of below if statments follows mathcad structure
        for text_piece in paragraph_elem:
            # Remove self.ws from tag for easier lookup
            current_emphasis_tag = text_piece.tag.replace(self.ws, "")
            current_emphasis_tag = current_emphasis_tag.replace(self.un, "")

            # Check if current tag is a emphasis formatting tag recognized in LaTeX
            if current_emphasis_tag in emphasis_latex:
                # Add start of latex formatting tag. Remove self.ws to lookup in our dict
                text += emphasis_latex[current_emphasis_tag]

            # Check if there's text to grab, if yes, add it to the string
            if text_piece.text is not None:

                # Check text is not only tab charecters
                if "\t" not in text_piece.text:
                    text += text_piece.text

            # If the tag is "region" then we have found a math region
            if current_emphasis_tag == "region":

                # Call math_reader to format the math
                text += " $ " + self.math_reader(text_piece[0][0]) + " $ "

                # Add tailing text, if it exists, not if it's just tab (\t)
                if (text_piece.tail is not None) and ("\t" not in text_piece.tail):
                    text += text_piece.tail

                # Skip sibling tags since they are handled with above code
                continue

            # Add <sp>(ace) charecters in mid-sentences specified by attribute "count" for tag
            if current_emphasis_tag == "sp":

                # Check if "count" attribute (key) exists
                if "count" in text_piece.attrib:
                    number_of_spaces = int(text_piece.attrib["count"])

                # If there's no attribute, there's only one space
                else:
                    number_of_spaces = 1

                text += " " * number_of_spaces

            # Go deeper if possible
            # If theres children, there's no text to grap, we need to go a level deeper to find text
            # (Check if text_piece has children which occurs using nested-formatting (eg __**bold&underlined**__))
            # The "f" tag is used for specificng fonts in mathcad, but we don't care about font went exporting to LaTeX
            if bool(text_piece):
                # Go one level deeper in the tree
                text += self.text_piece_formatter(text_piece)

            # Add the last part of LaTeX formatting
            if current_emphasis_tag in emphasis_latex:
                # Add text and end of latex formatting tag }
                text += emphasis_latex[current_emphasis_tag]

            # Add tailing text, if it exists, not if it's just tab (\t)
            if (text_piece.tail is not None) and ("\t" not in text_piece.tail):
                text += text_piece.tail

        return text

    def picture_reader(self, elem):
        """Method for reading binary picture data

        This method converts the desired picture from base64 data to a .png image
        :param elem: The picture region as a ElementTree
        :return: The LaTeX code including the image file
        """
        image_id = int(elem[0].attrib["item-idref"])  # Grap the image ID from the elements attributes
        image_base64_data = self.math_tree[4][image_id - 1].text  # Find the image data in the binaryContent part
        image_base64_data = image_base64_data.encode(encoding='UTF-8')  # Encode string as bytes instead
        filename = self.filename + "_" + str(image_id) + ".png"
        filename_no_ext = self.filename + "_" + str(image_id)

        # Open a new file for writing, name it the filename_id
        with open(self.output_folder + "/" + filename, "wb") as imagefile:  # Open the file, and close it afterwards
            imagefile.write(base64.decodebytes(image_base64_data))  # Write the decoded base64 bytes to the file

        return "\\includegraphics{\"" + filename_no_ext + "\"}"

    def latex_formatter(self, operator, x, y=None):  # Define the value of y
        """LaTeX math syntax formatter metod

        This method takes in the data math_reader() has found in the XML file,
        and formats it into LaTeX expressions.

        :param operator: A math operator or similar
        :param x: The first part of the expression
        :param y: The last part of the expression - Default value: None
        :return: Math expression formatted in LaTeX format
        """
        operator = operator.replace(self.ml, "")  # Remove prefix from operator
        operator = operator.replace(self.un, "")

        operators_mathcad_tag_to_latex = {
            'plus': "{x} + {y}",
            'minus': "{x} - {y}",
            'mult': "{x}·{y}",
            'div': "{x}/{y}",
            'eval': '{x} = {y}',
            'equal': '{x} = {y}',
            'define': '{x} = {y}',
            'symEval': '{x} = {y}',
            'globalDefine': '{x} = {y}',
            'pow': '{x}^{y}^',
            'nthRoot': "{x}^1/{y}^",
            'lessThan': '{x} < {y}',
            'greaterThan': '{x} > {y}',
            'lessOrEqual': '{x} \u2264 {y}',
            'greaterOrEqual': '{x} \u2265 {y}',
            'and': '{x} \\land {y}',
            'or': '{x} \\lor {y}',
            'apply': '{x}({y})',  # Return two strings that needs to be merged ('{x}\\left({y}\\right)')
            'function': '{x}({y})', # '{x}\\left({y}\\right)'
            'indexer': '{x}_{{{y}}}',
            'parens': '({x})',
            'sqrt': '\u221A({x})',  # Latex output: \sqrt{x} double {{ to esc
            'absval': '\\left|{x}\\right|',
            'neg': '-{x}',
            'vectorize': '\\vec{{{x}}}',
            'unitedValue': '{x}·{y}',
            'complex': '({x} + {y})'
        }

        if y is not None:  # y exists, so there's two components to parse
            if self.debug: print("y given")

            if operator in operators_mathcad_tag_to_latex:
                return operators_mathcad_tag_to_latex[operator].format(x=x, y=y)

            elif operator == "matrix":
                string = "\\begin{pmatrix}\n"
                i2 = 1
                # Rows and cols in the matrix, taken from the y tuple
                rows = y[0]
                cols = y[1]

                # We have 2 counters; one to keep track of current row (i), one for col (i2)
                for i in range(0, rows):  # Run this loop the amount of rows there exists
                    for entity in x[i, :]:  # For each value in the i'nte row
                        if i2 == cols:  # RNS checkmate
                            string += entity
                        else:
                            string = string + entity + " & "
                        i2 += 1
                    i2 = 1
                    string += "\\\\\n"

                string += "\\end{pmatrix}"
                return string

            elif operator == "integral":  # For integrals with limits
                lim_a = self.math_reader(y[0])
                lim_b = self.math_reader(y[1])
                var = self.math_reader(x[0])
                func = self.math_reader(x[1])
                return "\\int_{" + lim_a + "}^{" + lim_b + "} " + func + " d" + var

            elif operator == "derivative":  # For n'te derivative notation
                return "\\frac{d^" + y + "}{d" + self.math_reader(x[0]) + "^" + y + "}" + self.math_reader(x[1])

            else:
                return "Unhandled tag (y given) :("

        else:  # Else, there is only 1 value
            if self.debug: print("No y given")

            if operator in operators_mathcad_tag_to_latex:
                return operators_mathcad_tag_to_latex[operator].format(x=x)

            if operator == "id":
                # Make sure subscript attribute exists

                if x.get("subscript") is not None:
                    upper_text = symbol_parser(x.text, True)

                    # Get attribute's values and send to symbol_parser
                    sub_text = symbol_parser(x.attrib["subscript"], True)

                    return f"{upper_text}|{sub_text}|"
                #
                #     return upper_text + "_{" + sub_text + "}"
                else:
                    return symbol_parser(x.text, True)

            elif operator == "derivative":  # For derivative notation
                return "\\frac{d}{d" + self.math_reader(x[0]) + "}" + self.math_reader(x[1])

            elif operator == "integral":  # For integrals.
                var = self.math_reader(x[0])
                func = self.math_reader(x[1])
                return "\\int " + func + " d" + var

            else:
                return "Unhandled tag (y given) :("

    def main(self):
        """Method for controlling file writing

        Writes the final results to the new .tex file
        """
        self.tex_file.write(self.start_latex_doc)  # Write start of LaTeX document

        for child in self.math_tree[3]:  # Run for each region containing math or text
            self.current_region_no += 1  # Update counter
            print("\nTrying to parse the " + str(self.current_region_no) + "' region")  # Line separator for ouput
            # ToDo: Print the actual region-id
            # print("(region-id: " + child[0].attr["region-id"] + ")")

            try:  # Try to parse
                # ToDo: Smart align, that check for next region, if it's text or not?
                if child[0].tag == self.ws + "math":  # Math region
                    if self.debug: print("Type: Math region.")
                    # Write result of the region by calling fuction which sends the current element
                    string = self.math_reader(child[0][0])

                    # The part for substituting variable values
                    if self.auto_substitute:
                        pos1 = string.find('=')  # Find the first =
                        if pos1 != -1:
                            start = pos1 + 1
                            pos2 = string.find('=', start)  # Find the second =
                            if pos2 != -1:
                                start = string[:pos1]
                                defin = string[pos1:pos2]
                                substit = '' # Make an empty substit for now
                                for key in self.variables.keys(): # Check if defin has any variables
                                    if defin.find(key) != -1:
                                        substit += defin # If found a variable name create a substit
                                        break # Else it stays empty
                                end = string[pos2:]
                                # Dictionary is unordered, so when we replace elements with short names like "L" or
                                # "K" there's a chance they'll replace a part of a longer name. So they have to be
                                # replaced last. We need to sort the names by length in desc order
                                list_dict = [] # Make a temporary list
                                for key in self.variables.keys(): # Rewrite the variable dictionary to it
                                    list_dict.append([key, self.variables[key]])
                                list_dict.sort(key = lambda x:len(x[0]), reverse = True) # Sort the list
                                for variable in list_dict:
                                    # Replace variable names with their values from the sorted list
                                    substit = substit.replace(variable[0], variable[1])
                                string = start + defin + substit + end

                    self.tex_file.write(string + '\n')

                elif child[0].tag == self.ws + "text":  # Handle pure text regions
                    if self.debug: print("Type: Text region.")
                    self.tex_file.write(self.text_reader(child[0])+'\n')

                elif child[0].tag == self.ws + "picture":  # Handle pure picture regions
                    if self.debug: print("Type: Picture region.")
                    self.tex_file.write(self.picture_reader(child[0]) + "\\\\\n")

            # Catch the most common error
            except TypeError as e:
                print("Unsupported expessions found OR error occured, could not parse region", self.current_region_no,
                      "Error:", str(e))

        self.tex_file.write(self.end_latex_doc)  # Write end of LaTeX document

root_widget = Tk()  # Initialize Tkinter by creating a root widget others will be children of
new_app = ParseGUI(root_widget)  # Create a new instance of the ParseGUI class using the root widget as a parent

root_widget.wm_title("Jimmie's Wonderland")  # Set window title
root_widget.geometry('380x160')  # Set window size
root_widget.resizable(width=False, height=False)  # Make the window non-resizable
root_widget.mainloop()  # Start a loop that ends when the quit event is called (X button on window)