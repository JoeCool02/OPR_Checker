#!/usr/bin/env python

"""***Windows spell-check program in Python:***

import win32com.client

msword = win32com.client.Dispatch("Word.Application")

string = "Here is some texxt with a coupple of problems"

for word in string.split():
    if msword.CheckSpelling(word):
       print "!%s! OK!" % word
    else:
       print "?%s? ->"%word,
       suggestions = msword.GetSpellingSuggestions(word)
       for suggest in suggestions:
           print suggest.Name,
       print
"""

"""***Find Double Words***
p = re.compile(r'(\b\w+)\s+\1')
p.search('Paris in the the spring').group()
'the the'
"""

import os, popen2
import xml.dom.minidom
import re
import gzip, zipfile

class pr_object:
    
    def __init__(self, pr_filename, settings_filename):
        #Initialize Spell Checker
        self.spell_check = ispell()
   
        #Load PR and list of settings
        print "Converting XFDL to XML..."
        
        #Load PR into parsed XML document
        os.system('uudeview -i %s' % pr_filename)
        zipped_file_path = os.path.split(pr_filename)[0]
        zipped_file = '%s%sUNKNOWN.001'% (zipped_file_path, os.path.sep)
        zip_data = gzip.GzipFile(zipped_file,'r')
        self.doc = xml.dom.minidom.parse(zip_data)
        zip_data.close()
        os.remove(zipped_file)

        #Load settings into parsed XML document
        OD_TABLE_NS = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'
        zip_data = zipfile.ZipFile(settings_filename)
        content = zip_data.read('content.xml')
        self.settings = xml.dom.minidom.parseString(content)
        zip_data.close()

        #Load REGEX checks into dictionary
        sheet_list = self.settings.getElementsByTagNameNS(OD_TABLE_NS, 'table')

        for sheet in sheet_list:
            sheet_name = sheet.getAttributeNS(OD_TABLE_NS, 'name')
            sheet_rows = sheet.getElementsByTagNameNS(OD_TABLE_NS, 'table-row')                

            if sheet_name == "OPR Fields":
                 field_dict_p1, field_dict_p2 = self.get_cells(sheet_rows, OD_TABLE_NS)

            if sheet_name == "OPR Checks":
                 checks_dict_p1, checks_dict_p2 = self.get_cells(sheet_rows, OD_TABLE_NS)

            if sheet_name == "OPR Popups":
                 popups_dict_p1, popups_dict_p2 = self.get_cells(sheet_rows, OD_TABLE_NS)
                      
        print "Page 1 Field Dictionary"
        self.print_dict(field_dict_p1)
        print "Page 2 Field Dictionary"
        self.print_dict(field_dict_p2)

        print "Page 1 Check Box Dictionary"
        self.print_dict(checks_dict_p1)
        print "Page 2 Check Box Dictionary"
        self.print_dict(checks_dict_p2)

        print "Page 1 Popups Dictionary"
        self.print_dict(popups_dict_p1)
        print "Page 2 Popups Dictionary"
        self.print_dict(popups_dict_p2)

        #Store globalpage, page 1 & page 2 in attributes
        self.globalpage = self.doc.getElementsByTagName('globalpage')[0]
        pages = self.doc.getElementsByTagName('page')
        self.page1 = pages[0]
        self.page2 = pages[1]
        self.main(field_dict_p1, field_dict_p2, checks_dict_p1, checks_dict_p2, \
                  popups_dict_p1, popups_dict_p2)

    def get_cells(self, sheet_rows, OD_TABLE_NS):
        
        dict_p1 = {}
        dict_p2 = {}

        for row in sheet_rows:
            cells = row.getElementsByTagNameNS(OD_TABLE_NS, 'table-cell')
            if self.get_text(cells[4]) == "1":         
                dict_p1[self.get_text(cells[1])] = (self.get_text(cells[0]), \
                                                    self.get_text(cells[2]), \
                                                    self.get_text(cells[3]), \
                                                    self.get_text(cells[4]), \
                                                    self.get_text(cells[5]) \
                                                    )
            elif self.get_text(cells[4]) == "2":
                dict_p2[self.get_text(cells[1])] = (self.get_text(cells[0]), \
                                                    self.get_text(cells[2]), \
                                                    self.get_text(cells[3]), \
                                                    self.get_text(cells[4]), \
                                                    self.get_text(cells[5]) \
                                                    )
        return dict_p1, dict_p2

    def get_text(self, node):
        text = ''
        for child in node.childNodes:
            if child.nodeType == child.ELEMENT_NODE:
                text = text+self.get_text(child)
            elif child.nodeType == child.TEXT_NODE:
                text = text+child.nodeValue
        return text.encode('ascii')
    
    def test_group(self, check_tuple, check_type):

        p1_on_form = self.get_on_form_dicts(check_type, self.page1)
        p2_on_form = self.get_on_form_dicts(check_type, self.page2)
        
        on_form = (p1_on_form, p2_on_form)       

        #Spell Check Applicable Fields

        print "Running Spell Check..."

        for page in range(2):
            for item in on_form[page]:
                try:
                    if check_tuple[page][item][4] == 'Y':
                        for i in on_form[page][item].split():
                            print i
                            result = self.spell_check(i)
                            if result != None:
                                print "Field: ", check_tuple[page][item][0], "=> Spell Check Failed"
                                print "Found: ", i
                                print "Possible Alternatives: ", result, "\n"
                            else:
                                print result
                except:
                    pass

        print "\n"
   
        for i in range(2):
            self.regex_check(on_form[i], check_tuple[i])
    
    def get_on_form_dicts(self, form_field, xmldoc):
        elements = xmldoc.getElementsByTagName(form_field)
        form_dict = self.readxmlnodes(elements)
        return form_dict
    
    def print_dict(self, the_dict):
        for i in the_dict:
            print i, the_dict[i]
    
    def readxmlnodes(self, xmllist):

        value_dict = {}

        for i in xmllist:
            j = i.getElementsByTagName('value')
            try:
                value_dict[j[0].parentNode.getAttribute('sid')] = j[0].firstChild.nodeValue
            except:
                value_dict[j[0].parentNode.getAttribute('sid')] = None
        return value_dict
    
    def regex_check(self, IUT_dict, truth_dict):

       print "Running Regular Expression Check..."

       for i in IUT_dict:
             try:
                 if truth_dict[i][2] != 'None':
                     pattern = re.compile(truth_dict[i][2])
                     try:
                         match = pattern.match(IUT_dict[i]).group()
                         print truth_dict[i][0], "=> [OK]\n", "Field:", i, "Text:", match, "\n"
                     except:
                         print truth_dict[i][0], "=> [FAIL]\n", "Field:", i, "Text:", \
                         IUT_dict[i], "\n"
                         self.fails += 1
             except KeyError:
                 pass

    #def spell_check
    #def hyphen_check
    #def no_hyphen_check
    #def sig_block_check
    #def get_hyphen_words
    #def get_non_hyphen_words
    #def get_acros_and_abbrevs
    
    def main(self, *args):


        print "\n*****************START PR ANALYSIS*******************\n"
 
        self.fails = 0

        self.test_group((args[0], args[1]), 'field')
        self.test_group((args[2], args[3]), 'check')
        self.test_group((args[4], args[5]), 'popup')

        print "There are %d failed fields in this PR." % self.fails

class ispell:
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

def usage():
    print "Usage: prchecker filename"


if __name__ == "__main__":
    p = optparse.OptionParser()
    options, arguments = p.parse_args()
    if len(arguments)!=1:
         usage()
    pr_file = """%s%s%s""" % (os.getcwd(), os.path.sep, arguments[0])
    try:
        settings_file = """%s%sPR Structure.ods""" % (os.getcwd(), os.path.sep)
    except:
        print "PR Structure.ods is missing.  Please insure it is in the same directory \
               as the prchecker executable."

    print pr_file, settings_file

    #pr = pr_object(pr_file, settings_file)
    exit(0)



