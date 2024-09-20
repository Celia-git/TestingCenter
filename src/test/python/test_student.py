import unittest

from .__init__ import *


class StudentTester(unittest.TestCase):

    def __init__(self, methodName: str = "runTest") -> None:
        super().__init__(methodName)
        self.student = Student(data_path, "A00123456")
    
    def test_get_aNum(self):
        self.assertEqual(self.student.get_aNum(), "A00123456")

    def test_get_name(self):
        self.assertEqual(self.student.get_name(), "Finley Jones")
    
    def test_get_courses(self):
        self.assertEqual(self.student.get_courses(), [['2110', 'N30']])
    
    def test_get_visits(self):
        self.assertEqual(self.student.get_visits(), ['04/15/24'])

    def test_get_save_data(self):
        self.assertEqual(self.student.save_data(), 0)

    def test_has_course(self):
        self.assertEqual(self.student.has_course(['2110', 'N30']), True)

    def test_load_data(self):
        new_student = Student(data_path, "A00898989")
        self.assertEqual(new_student.aNum, "A00898989")
        self.assertEqual(new_student.name, "Arely Barry")
        self.assertEqual(new_student.courses, [["1710", "N35"]])
        self.assertEqual(new_student.visits, ["04/07/24"])

def run_tests():
    unittest.main(argv=['first-arg-is-ignored'], exit=False)