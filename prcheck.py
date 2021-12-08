#!/usr/bin/env python

"""Author: Capt Josef Peterson
(2009) All Rights Reserved

Version: 0.5.1

Purpose:  Identify and report common errors in Air Force performance reports.

Tested with:  	AF707, 2008/06/18, ver 2.79.9
              	AF910, 2006/12/01, ver 8.113.5

Known Issues:
 
-Will not open IE with Adobe Acrobat if current PR output file is already open
-Does not include checking for words that need hyphens or vice versa
-Does not identify the field or line where spelling errors are found
-Does not check for the correct form version and exit gracefully
-Does not check for valid abbreviations and acronyms
-Fails duty title containing "ERS CC"

Nice To Haves:
-Explanations for check fails
-PRF Checking

Static Files:
-PR Structure.ods: Contains regex and check suite information for PRs
-Text2PDF.exe (Windows) or Text2PDF (Linux): Converts raw text output to PDFs (Thanks to Anand B Pillai)
-setup.py: Configuration information for py2exe
-uudeview.exe (Windows) or uudeview (Linux): Converts .XFDL file to gzipped XML
-prchecker_splash.gif: Image for the splash
-others (may be included with py2exe distribution)

Dynamic Files:
-logfile.txt: Containes stdout feed
-logfileerr.txt: Containes stderr feed
-output.pdf: Contains the output from the PR Checker
"""

import sys

sys.stdout = open("logfile.txt", "w")
sys.stderr = open("logfileerr.txt", "w")
import os, shutil
import xml.dom.minidom
import re
import gzip, zipfile

try:
    import win32gui
    import win32com.client
except:
    pass
import optparse
import Tkinter
from tkFileDialog import askopenfilename


class pr_object:
    def __init__(self, pr_filename, settings_filename):

        # Initialize Output File
        self.pr_filename = pr_filename
        self.output_filename = "%s%s%s.out" % (
            os.getcwd(),
            os.path.sep,
            os.path.basename(pr_filename),
        )
        program_string = """=====PR Checker=====\n\n"""
        author_string = (
            """Author: Capt Josef Peterson\n(2009) All Rights Reserved\n\n"""
        )
        file_name_string = """File: %s\n\n""" % os.path.basename(pr_filename)
        self.output = open(self.output_filename, "w", -1)
        self.output.write(program_string)
        self.output.write(author_string)
        self.output.write(file_name_string)

        # Load PR into parsed XML document
        print("Converting XFDL to XML...")
        if os.name == "posix":
            os.system("""uudeview -i "%s" """ % pr_filename)
        else:
            os.system(
                """%s%suudeview -i "%s" """ % (os.getcwd(), os.path.sep, pr_filename)
            )
        zipped_file = "%s%sUNKNOWN.001" % (os.getcwd(), os.path.sep)

        try:
            zip_data = gzip.GzipFile(zipped_file, "r")
            self.doc = xml.dom.minidom.parse(zip_data)
        except:
            print(
                "Cannot convert file.  It may be an outdated PR version.  Contact your administrator."
            )
            self.output.write(
                "Cannot convert file.  It may be an outdated PR version.  \
                               Contact your administrator.\n"
            )
            self.clean_up()

        finally:
            zip_data.close()
            os.remove(zipped_file)

        # Store globalpage, page 1 & page 2 in attributes
        self.globalpage = self.doc.getElementsByTagName("globalpage")[0]
        pages = self.doc.getElementsByTagName("page")
        self.page1 = pages[0]
        self.page2 = pages[1]

        # Determine and write PR type, version
        pr_types = {"Officer": "OPR", "ENLISTED": "EPR"}
        self.pr_type_text = self.get_text(
            self.globalpage.getElementsByTagName("title")[0]
        )
        self.pr_version_text = self.get_text(
            self.globalpage.getElementsByTagName("custom:date")[0]
        )
        self.output.write("Type: ")
        self.output.write(self.pr_type_text)
        self.output.write("\n\n")
        self.output.write("Version: ")
        self.output.write(self.pr_version_text)
        self.output.write("\n\n")

        for pr_type in pr_types:
            p = re.compile(pr_type)
            if p.search(self.pr_type_text):
                self.pr_type = pr_types[pr_type]

        # Load settings into parsed XML document
        OD_TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
        zip_data = zipfile.ZipFile(settings_filename)
        content = zip_data.read("content.xml")
        self.settings = xml.dom.minidom.parseString(content)
        zip_data.close()

        # =Load REGEX checks into dictionaries=
        sheet_list = self.settings.getElementsByTagNameNS(OD_TABLE_NS, "table")

        # Pull cells from tables corresponding to PR type
        fields_str = "%s Fields" % self.pr_type
        checks_str = "%s Checks" % self.pr_type
        popups_str = "%s Popups" % self.pr_type

        for sheet in sheet_list:
            sheet_name = sheet.getAttributeNS(OD_TABLE_NS, "name")
            sheet_rows = sheet.getElementsByTagNameNS(OD_TABLE_NS, "table-row")

            if sheet_name == fields_str:
                field_dict_p1, field_dict_p2 = self.get_cells(sheet_rows, OD_TABLE_NS)

            if sheet_name == checks_str:
                checks_dict_p1, checks_dict_p2 = self.get_cells(sheet_rows, OD_TABLE_NS)

            if sheet_name == popups_str:
                popups_dict_p1, popups_dict_p2 = self.get_cells(sheet_rows, OD_TABLE_NS)

            if sheet_name == "Senior Rater Info":
                self.SR_dict = {}
                for row in sheet_rows:
                    cells = row.getElementsByTagNameNS(OD_TABLE_NS, "table-cell")
                    self.SR_dict[self.get_text(cells[1])] = (
                        self.get_text(cells[0]),
                        self.get_text(cells[2]),
                    )

            if sheet_name == "PR Version":
                self.ver_dict = {}
                for row in sheet_rows:
                    cells = row.getElementsByTagNameNS(OD_TABLE_NS, "table-cell")
                    self.ver_dict[self.get_text(cells[0])] = self.get_text(cells[1])

            if sheet_name == "Overlook":
                self.overlook_list = []
                cells = sheet.getElementsByTagNameNS(OD_TABLE_NS, "table-cell")
                for cell in cells:
                    if self.get_text(cell):
                        self.overlook_list.append(self.get_text(cell))
                if options.verbose:
                    self.output.write(repr(self.overlook_list) + "\n")

            if sheet_name == "Catch":
                self.catch_list = []
                cells = sheet.getElementsByTagNameNS(OD_TABLE_NS, "table-cell")
                for cell in cells:
                    if self.get_text(cell):
                        self.catch_list.append(self.get_text(cell))
                if options.verbose:
                    self.output.write(repr(self.catch_list) + "\n")

        # Print the contents of the dictionaries if the verbose option is set
        if options.verbose:

            self.print_dict(field_dict_p1, "Page 1 Field Dictionary")
            self.print_dict(field_dict_p2, "Page 2 Field Dictionary")

            self.print_dict(checks_dict_p1, "Page 1 Check Box Dictionary")
            self.print_dict(checks_dict_p2, "Page 2 Check Box Dictionary")

            self.print_dict(popups_dict_p1, "Page 1 Popups Dictionary")
            self.print_dict(popups_dict_p2, "Page 2 Popups Dictionary")

            self.print_dict(self.SR_dict, "Senior Rater Info")
            self.print_dict(self.ver_dict, "Version Information")

        # Initialize Spell Checker
        if os.name == "nt":
            self.spell_checker = msword_spell_check(self.overlook_list)
            print("Spell Checker Initialized")

        # Start program main function
        self.main(
            field_dict_p1,
            field_dict_p2,
            checks_dict_p1,
            checks_dict_p2,
            popups_dict_p1,
            popups_dict_p2,
        )

    def get_cells(self, sheet_rows, OD_TABLE_NS):
        # Pull the fields out of the ODS spreadsheet and put in dictionaries.
        dict_p1 = {}
        dict_p2 = {}

        for row in sheet_rows:
            cells = row.getElementsByTagNameNS(OD_TABLE_NS, "table-cell")
            try:
                if self.get_text(cells[4]) == "1":
                    dict_p1[self.get_text(cells[1])] = (
                        self.get_text(cells[0]),
                        self.get_text(cells[2]),
                        self.get_text(cells[3]),
                        self.get_text(cells[4]),
                        self.get_text(cells[5]),
                    )
                elif self.get_text(cells[4]) == "2":
                    dict_p2[self.get_text(cells[1])] = (
                        self.get_text(cells[0]),
                        self.get_text(cells[2]),
                        self.get_text(cells[3]),
                        self.get_text(cells[4]),
                        self.get_text(cells[5]),
                    )
            except:
                pass

        return dict_p1, dict_p2

    def get_text(self, node):
        # Pull text out of one XML node
        text = ""
        for child in node.childNodes:
            if child.nodeType == child.ELEMENT_NODE:
                text = text + self.get_text(child)
            elif child.nodeType == child.TEXT_NODE:
                text = text + child.nodeValue
        return text.encode("ascii")

    def get_on_form_dicts(self, form_field, xmldoc):
        # Return a dictionary with all keys & items from elements matching the string form_field
        # from an XML document
        xmllist = xmldoc.getElementsByTagName(form_field)
        value_dict = {}

        for i in xmllist:
            j = i.getElementsByTagName("value")
            try:
                value_dict[j[0].parentNode.getAttribute("sid")] = j[
                    0
                ].firstChild.nodeValue
            except:
                value_dict[j[0].parentNode.getAttribute("sid")] = ""

        return value_dict

    def print_dict(self, the_dict, name):
        # Print the keys and items in an arbitrary dictionary.
        if the_dict and name:
            print(name)
            self.output.write(name)
            self.output.write("\n")
            for i in the_dict:
                print(i, the_dict[i])
                self.output.write(i)
                self.output.write(" ")
                self.output.write(repr(the_dict[i]))
                self.output.write("\n")
            self.output.write("\n")

    def regex_check(self, IUT_dict, truth_dict):
        # Check in incoming dictionary against regular expressions in a truth dictionary
        if len(IUT_dict) == 0:
            return

        print("Running Regular Expression Check...\n")

        for i in IUT_dict:
            try:
                if truth_dict[i][2] != "None":
                    pattern = re.compile(truth_dict[i][2])
                    try:
                        match = pattern.match(IUT_dict[i]).group()
                        if options.verbose:
                            out_string = "%s => [OK]\nField: %s Text: %s\n" % (
                                truth_dict[i][0],
                                i,
                                match,
                            )
                            print(out_string)
                            self.output.write(out_string)
                            self.output.write("\n")
                    except:
                        out_string = "%s => [FAIL]\nField: %s Text: %s\n" % (
                            truth_dict[i][0],
                            i,
                            IUT_dict[i].encode("unicode_escape"),
                        )
                        print(out_string)
                        self.output.write(out_string)
                        self.output.write("\n")
                        self.fails += 1
            except KeyError:
                pass

    def spell_check(self, on_form_dict, check_dict):
        # Use the msword spelling object to spell check fields
        print("Running Spell Check...")

        for item in on_form_dict:
            try:
                spell_item = check_dict[item][4]
                if spell_item == "Y":
                    lines = on_form_dict[item].splitlines()
                    line_count = 1
                    for line in lines:
                        spelling_result, warnings = self.spell_checker(line)
                        self.warnings += warnings
                        if spelling_result:
                            line_out = (
                                "\n"
                                + check_dict[item][0]
                                + " "
                                + "line %d" % line_count
                                + ":\n"
                            )
                            self.output.write(line_out)
                            self.output.write(spelling_result)
                        line_count += 1
            except KeyError:
                pass

    def version_check(self, read_version, version_dict):
        # Check if the form is the right version -- use information form PR Structure.ods
        if version_dict[self.pr_type] != read_version:
            self.output.write("Version Check => [FAIL]\n")
            self.output.write("Correct PR Version: %s\n" % version_dict[self.pr_type])
            self.output.write("This PR Version: %s\n\n" % read_version)
        elif options.verbose:
            self.output.write("Version Check => [OK]\n\n")

    def catch_common(self, on_form_dict, check_dict):
        # Look for common error patterns defined on Catch sheet in PR Structure.ods
        for item in on_form_dict:
            try:
                if check_dict[item][4] == "Y":
                    line_number = 1
                    for line in on_form_dict[item].splitlines():
                        for pattern in self.catch_list:
                            p = re.compile(r"%s" % pattern)
                            match = p.search(line)
                            if match:
                                output = (
                                    "\n%s, Line %d:\n[WARNING] Likely error => %s\n\n"
                                    % (check_dict[item][0], line_number, match.group())
                                )
                                self.output.write(output)
                                self.warnings += 1
                            else:
                                if options.verbose:
                                    self.output.write(
                                        "Catch Common => %s [OK]\n" % pattern
                                    )
                        line_number += 1

            except KeyError:
                pass

    # def senior_rater_sig_block_match

    def clean_up(self):
        # Finish up the script and call the output files
        self.output.close()

        os.system(
            '%s%sText2PDF "%s"' % (os.getcwd(), os.path.sep, self.output_filename)
        )
        self.pdfout_file = "%s%s%s.pdf" % (
            os.path.split(self.pr_filename)[0],
            os.path.sep,
            os.path.basename(self.pr_filename),
        )
        print(self.pdfout_file)
        print(self.output_filename)

        shutil.move("%s.pdf" % (self.output_filename), self.pdfout_file)
        os.remove(self.output_filename)

        if os.name == "posix":
            try:
                os.system('evince "%s"' % self.pdfout_file)
            except:
                pass
        else:
            ie = win32com.client.Dispatch("internetexplorer.application")
            ie.visible = 1
            ie.Navigate("%s" % self.pdfout_file)

    def test_group(self, check_tuple, check_type):
        # Test the form's front and back page against truth dictionaries
        p1_on_form = self.get_on_form_dicts(check_type, self.page1)
        p2_on_form = self.get_on_form_dicts(check_type, self.page2)

        on_form = (p1_on_form, p2_on_form)

        # Spell Check Applicable Fields & Check The Pages Against Regular Expressions
        for i in range(2):
            if os.name == "nt":
                self.spell_check(on_form[i], check_tuple[i])
            self.catch_common(on_form[i], check_tuple[i])
            self.regex_check(on_form[i], check_tuple[i])

    def main(self, *args):

        print("\n*****************START PR ANALYSIS*******************\n")

        please_wait = Tkinter.Tk()
        please_wait.title = "PR Checker"
        centerx = please_wait.winfo_screenwidth() / 2
        centery = please_wait.winfo_screenheight() / 2
        dimensions = (200, 25, centerx - 100, centery - 12)
        please_wait.geometry("%dx%d+%d+%d" % dimensions)
        wait_message = "Analyzing, Please Wait..."
        msg = Tkinter.Message(please_wait, text=wait_message, width=175)
        msg.pack()
        please_wait.update()

        self.fails = 0
        self.warnings = 0

        self.version_check(self.pr_version_text, self.ver_dict)

        self.test_group((args[0], args[1]), "field")
        self.test_group((args[2], args[3]), "check")
        self.test_group((args[4], args[5]), "popup")

        warning_string = "===%d warning(s)===\n" % self.warnings
        fail_string = "***%d failed field(s)***\n" % self.fails

        print(warning_string)
        print(fail_string)

        self.output.write(warning_string)
        self.output.write(fail_string)

        please_wait.destroy()

        self.clean_up()


# ****************SPELL CHECKER CLASSES*****************************


class msword_spell_check:
    def __init__(self, overlook_list):
        self.msword = win32com.client.Dispatch("Word.Application")
        self.msword.Documents.Add()
        self.overlook_res = []
        for word in overlook_list:
            self.overlook_res.append(re.compile(r"%s" % word, flags=re.I))

    def check_overlook(self, word):

        overlook = False

        for overlook_word in self.overlook_res:
            if overlook_word.match(word):
                overlook = True

        return overlook

    def __call__(self, string):
        warnings = 0
        output = ""
        for word in string.replace("-", " ").replace("/", " ").split():
            if self.msword.CheckSpelling(word):
                if options.verbose:
                    out_string = "!%s! OK!" % word
                    print(out_string)
                    output += out_string + "\n"
            else:
                if not self.check_overlook(word):
                    warnings += 1
                    out_string = "[WARNING] ?%s? ->" % word
                    print(out_string)
                    output += out_string
                    output += " "
                    suggestions = self.msword.GetSpellingSuggestions(word)
                    for suggest in suggestions:
                        out_string = suggest.Name
                        print(out_string)
                        output += out_string
                        output += " "
                    output += "\n"
                else:
                    if options.verbose:
                        out_string = "!%s! -> " % word + "matches overlook list.\n"
                        output += out_string

        return output, warnings


# ***********************GUI CLASSES********************************


class splash_image:
    def button_click_exit_mainloop(self, event):
        self.root.quit()  # this will cause mainloop to unblock.
        self.root.destroy()

    def __init__(self):
        self.root = Tkinter.Tk()
        self.gif = Tkinter.PhotoImage(file="prchecker_splash.gif")
        self.img = Tkinter.Label()
        self.img.pack()
        self.img.config(image=self.gif)
        self.button = Tkinter.Button(text="Process Performance Report")
        self.button.pack()
        self.root.bind("<Button>", self.button_click_exit_mainloop)
        self.centerx = (self.root.winfo_screenwidth() - self.gif.width()) / 2
        self.centery = (self.root.winfo_screenheight() - self.gif.height()) / 2
        self.dimensions = (
            self.gif.width(),
            self.gif.height() + 30,
            self.centerx,
            self.centery,
        )
        self.root.geometry("%dx%d+%d+%d" % self.dimensions)
        self.root.overrideredirect(1)
        self.root.tkraise()
        self.root.mainloop()


# ***********************START MAIN PROGRAM*************************


def find_file(title):
    if os.name == "posix":
        var = askopenfilename()
    elif os.name == "nt":
        var = win32gui.GetOpenFileNameW(Title=title)[0]
    else:
        print("Error: Couldn't Identify OS.  Consult Your Administrator")
        exit(0)
    return var


def usage():
    print("Usage: prchecker [options] filename")


if __name__ == "__main__":
    # Set up a couple of admin things to deal with windows' baloney
    working_dir = os.getcwd()

    splash_image()

    # Deal with arguments, options and incorrect usage
    p = optparse.OptionParser()
    p.add_option("--verbose", "-v", action="store_true")
    options, arguments = p.parse_args()
    if len(arguments) == 0:
        pr_file = ""
        while os.path.splitext(pr_file)[1] != ".xfdl":
            print("Please select input PR.")
            pr_file = find_file("Select Input PR")
            os.chdir(working_dir)
    elif len(arguments) > 1:
        usage()

    # Set up the input arguments for the PR Checker class; PR File & Settings File
    else:
        pr_file = """%s%s%s""" % (os.getcwd(), os.path.sep, arguments[0])
    settings_file = """%s%sPR Structure.ods""" % (os.getcwd(), os.path.sep)
    if not os.path.exists("""%s%sPR Structure.ods""" % (os.getcwd(), os.path.sep)):
        print("PR Structure.ods is missing.  Please find PR Structure.ods to continue.")
        settings_file = ""
        while os.path.splitext(settings_file)[1] != ".ods":
            print("Please select settings file.")
            settings_file = find_file("Select Settings File (PR Structure.ods)")
            os.chdir(working_dir)

    pr_file = os.path.abspath(pr_file)

    # Run the program, finally
    print("Analyzing PR...")
    pr = pr_object(pr_file, settings_file)

    # Clean up logfiles
    # sys.stdout.close()
    # sys.stderr.close()


# **********************NOTES AND EXTRAS***************************

"""class ispell:
    def __init__(self):
        self._f = popen2.Popen3("ispell -a")
        self._f.fromchild.readline() #skip the credit line
    def __call__(self, word):
        self._f.tochild.write(word+'\n')
        self._f.tochild.flush()
        s = self._f.fromchild.readline()
        if not (s[:1]=="*" or s[:1]=="+" or s[:1]=="-" or s[:1]=="#" or s[:1]=="?" or s[:1]=="&amp;"):
            return None
        #s = self._f.fromchild.readline()
        if s[:1]=="*" or s[:1]=="+" or s[:1]=="-":     #correct spelling
            return None
        elif s[:1]=="#":  # no matches
            return []
        else:
            m = re.compile("^[&amp;\?] \w+ [0-9]+ [0-9]+:([\w\- ,]+)$", re.M).search("\n"+s, 1)
            return (m.group(1).split(', '))
"""

"""***Find Double Words***
p = re.compile(r'(\b\w+)\s+\1')
p.search('Paris in the the spring').group()
'the the'
"""
