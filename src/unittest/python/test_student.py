import unittest

from __init__ import *


class StudentTester(unittest.TestCase):

    def __init__(self, methodName: str = "runTest") -> None:
        super().__init__(methodName)
        self.student = Student(data_path, "A00123456")
    
    def test_aNum(self):
        self.assertEqual(self.student.aNum, "A00123456")

    def test_name(self):
        self.assertEqual(self.student.name, "Finley Jones")
    
    def test_courses(self):
        self.assertEqual(self.student.courses, [['2110', 'N30']])
    
    def test_visits(self):
        self.assertEqual(self.student.visits, ['04/15/24'])

    def save_data(self):
        self.assertEqual(self.student.save_data(), 0)

    def test_has_course(self):
        self.assertEqual(self.student.has_course(['2110', 'N30']), True)

if __name__ == "__main__":
    unittest.main()