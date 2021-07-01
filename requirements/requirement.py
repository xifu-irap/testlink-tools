#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#     Copyright (c) Odile Coeur-Joly, IRAP Toulouse
#
#     This program is free software: you can redistribute it and/or modify
#     it under the terms of the GNU General Public License as published by
#     the Free Software Foundation, either version 3 of the License, or
#     (at your option) any later version.
#
#     This program is distributed in the hope that it will be useful,
#     but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#     GNU General Public License for more details.
#
#     You should have received a copy of the GNU General Public License
#     along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
#     requirement.py
#
from docx.opc.exceptions import PackageNotFoundError
from docx.api import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table, _Row
from docx.text.paragraph import Paragraph

"""This module manages the X-IFU/DRE requirements documents

1. Open Requirements.docx and read paragraphs and tables
2. Assert that :
    the doc is organized into chapters, (each chapter corresponds to a "requirement specification" for TestLink)
    the chapter header for requirements is of style = "Heading 2" and its title contains "requirements"
    every "requirement specification" contains one or more "requirement"
3. Convert from WORD to XML for TestLink ==> docx_to_XML (DONE)

4. Convert from WORD to EXCEL ==> .xlsx file (TODO)
5. convert from EXCEL to CSV ==> .csv file (TODO)
6. convert from CSV to XML ==> .xml file (TODO)
"""

class Requirement(object):
    """Manage Word (docx) requirement files
    """
    XML_DOC_START  = "<requirement-specification>"
    XML_SPEC_START = "   <req_spec>"
    XML_REQ_START  = "      <requirement>"
    XML_REQ_STOP   = "      </requirement>"
    XML_SPEC_STOP  = "   </req_spec>"
    XML_DOC_STOP   = "</requirement-specification>"

    def __init__(self, filename="", reqid="XIFU-DRE-DMX-FW-R", version="V1.0", level="SRS"):
        """Constructor
        """
        self.filename = filename
        self.xml_reqid = reqid
        self.xml_text = ['<?xml version="1.0" encoding="UTF-8"?>']
        self.xml_text.append('<requirement-specification>')
        
        # SPECIFICATION
        # type : 2=User Requirement Specification, 3 = System Requirement Specification
        self.xml_spec_tags = ['version', 'type', 'node_order', 'total_req']
        self.xml_spec_vals = [1, 3, 1, 0]
        self.xml_spec_dict = dict(zip(self.xml_spec_tags, self.xml_spec_vals))
        self.spec_doc_id = version + "-" + level
        
        # REQUIREMENT
        # status : D=Draft, R=Review, W=Rework, F=Finish, I= Implemented, V=Valid, N=Non Testable, O=Obsolete
        # type : 1=Informational, 2=Feature, 3=Use Case, 4=User Interface, 5=Non Functional, 6=Constraint, 7=System Function
        self.xml_req_tags = ['docid', 'title', 'version', 'revision', 'node_order', 'description', 'status', 'type', 'expected_coverage']
        self.xml_req_vals = [self.xml_reqid, 'title', 1, 1, 0, 'description', "V", 2, 0]
        self.xml_req_dict = dict(zip(self.xml_req_tags, self.xml_req_vals))
        
        # REQUIREMENT_TYPE
        # type : 1=Informational, 2=Feature, 3=Use Case, 4=User Interface, 5=Non Functional, 6=Constraint, 7=System Function
        

        try:
            self.document = Document(filename)
            print("document", self.document)
            
        except (ValueError, PackageNotFoundError) as error:
            print("ERROR {0} : {1}".format(type(error), error))

    def __iter_block_items(self, parent):
        """From : https://stackoverflow.com/questions/29240707/python-docx-get-tables-from-paragraph
        
        Generate a reference to each paragraph and table child within *parent*,
        in document order. Each returned value is an instance of either Table or
        Paragraph. *parent* would most commonly be a reference to a main
        Document object, but also works for a _Cell object, which itself can
        contain paragraphs and tables.
        """
        if isinstance(parent, _Document):
            parent_elm = parent.element.body
        elif isinstance(parent, _Cell):
            parent_elm = parent._tc
        elif isinstance(parent, _Row):
            parent_elm = parent._tr
        else:
            raise ValueError("something's not right")

        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    def __spec_to_xml(self, title, spec_id, scope):
        """Format a requirement specification to XML text
        """
        # Close the previous specification
        if spec_id > 1:
            self.xml_text.append(Requirement.XML_SPEC_STOP)
        
        # start to fill the XML specification
        self.xml_text.append('   <req_spec title=\"{0}\" doc_id=\"{1}\">'.format(title, self.spec_doc_id + str(spec_id)))
        for tag, val in self.xml_spec_dict.items():
            self.xml_text.append('      <{0}>{1}</{0}>'.format(tag, val))
        # Add the scope which is the next paragraph after the heading
        self.xml_text.append('      <{0}>{1}</{0}>'.format('scope', scope))

    def __req_to_xml(self, req_dict):
        """Format a requirement to XML text
        """
        self.xml_text.append(Requirement.XML_REQ_START)
        for tag, val in req_dict.items():
            self.xml_text.append('      <{0}>{1}</{0}>'.format(tag, val))
        self.xml_text.append(Requirement.XML_REQ_STOP)

    def docx_to_XML(self, document):
        """Scan a .docx document to extract Paragraphs and Tables for "requirement specification" and "requirements"
        """
        scope = ""
        heading_title = ""
        new_spec = False
        spec_id = 0
        
        # One pass to read the word document and fill the XML values
        for block in self.__iter_block_items(document):
            # read Paragraph
            if isinstance(block, Paragraph):
                # pickup the title of the paragraph as requirement specification title
                if block.style.name.lower().find("heading 2") >=0 and block.text.lower().find('requirements') >= 0:
                    heading_title = block.text
                    new_spec = True
                    continue
                else:
                    # append all lines between the header and the first table and build the scope
                    if new_spec:
                        scope = scope + block.text
                    continue
                # read table
            elif isinstance(block, Table):
                # the scope paragraph is ending just before the first table
                if new_spec:
                    spec_id = spec_id + 1
                    self.__spec_to_xml(heading_title, spec_id, scope)
                    new_spec = False
                    scope = ""
                # search for REQ table
                table = block
                for row in table.rows:
                    for cell in row.cells:
                        if cell.text.find(self.xml_reqid):
                            continue
                        else:
                            print("cell text=", cell.text)
                            for j, col in enumerate(table.columns):        
                                if j == 1:
                                    # for XML conversion: < and > are not allowed in text
                                    content = (cell.text.replace('<', 'lt').replace('>', 'gt') for cell in col.cells)
                                    l = list(content)
                                    # rearrange items from docx Document to XML for TestLink
                                    self.xml_req_vals[1] = l[0] # Title
                                    self.xml_req_vals[0] = l[1] # Reference REQ_ID
                                    self.xml_req_vals[5] = l[2] # Description
                                    self.xml_req_vals[7] = 2 # Type
                                    self.xml_req_vals[6] = str(l[4])[0] # Status
                                    req_dict = dict(zip(self.xml_req_tags, self.xml_req_vals))
                                    self.__req_to_xml(req_dict)

        # build the full XML text
        full_xml_text = '\n'.join(self.xml_text) + '\n' + Requirement.XML_SPEC_STOP +  '\n' + Requirement.XML_DOC_STOP
        return full_xml_text

if __name__ == '__main__':
    """main method to test this script as a unit test.
    """
    
    import easygui
    
    # Select input requirement (docx file)
    filename = easygui.fileopenbox("Please select a file", default="tests/*.docx")

    # Select the Document ID
    docid = easygui.enterbox("Requirement ID ?", default="DRE-DMX-FW-REQ")
    typespec = easygui.choicebox(msg = "Type of Specification", title = "Here", 
                             choices = ['Section','SRS','USR'], preselect=1)

    #convert to XML text for TestLink
    try:
        requirement = Requirement(filename, docid, "V0.8", typespec)
        xml = requirement.docx_to_XML(requirement.document)
    
        print("XML================\n", xml)
    
        f = open("result.xml", 'w', encoding='utf-8')
        f.write(xml)
        f.close()
    except:
        print("Please choose a valid .docx file...")
        
else:
    print("Importing...", __name__)
