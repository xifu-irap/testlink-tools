'''
Created on 24 juin 2021

@author: Odile
'''
import unittest
from docx.document import Document

from requirements.requirement import Requirement

class TestRequirement(unittest.TestCase):


    def test_requirement_constructor(self):
        
        req = Requirement("D:/MYCORE_OCJ/S.O.F.T.S/ATHENA/PYTHON/requirements/data/0065-Requirements.docx", "V3.0", "User")
        self.assertIsInstance(req.document, Document)

if __name__ == "__main__":
    #import sys;sys.argv = ['', 'Test.testName']
    unittest.main()