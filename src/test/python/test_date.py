

from .__init__ import *

class DateTester(unittest.TestCase):


    def __init__(self, methodName: str = "runTest") -> None:
        super().__init__(methodName)
        self.date = Date(data_path)



    def test_load_data(self):
        data = self.date.load_data("05/29/24", 10)
        self.assertEqual(list(data.keys())[0], "05/29/24")
        self.assertEqual(list(data.values())[0][0][0], 'A00223332')
        self.assertEqual(list(data.values())[0][0][1], 'FALSE')
        self.assertEqual(list(data.values())[0][0][2], '1130')
        self.assertEqual(list(data.values())[0][0][3], 'N30')
        self.assertEqual(list(data.values())[0][0][4], '55')
        self.assertEqual(list(data.values())[0][0][5], '14:53')
        self.assertEqual(list(data.values())[0][0][6], '18:13')

    def test_get_index(self):
        wb = openpyxl.load_workbook(data_path, read_only=True)
        ws = wb["Date"]
        [start_row, end_row] = self.date.get_index("04/19/24", 5, ws)
        self.assertEqual(start_row, 10)
        self.assertEqual(end_row, 8)

    def test_get_int_date(self):
        self.assertEqual(self.date.get_int_date("06/29/24"), (179, 24))

    def test_save_data(self):
        data= self.date.load_data("05/29/24", 10)
        self.assertEqual(self.date.save_data(data), 0)

    def test_save_student_checkout(self):
        result = self.date.save_student_checkout('05/29/24', 'A00223332', '14:53', '18:13')
        self.assertEqual(result, (0, 0))
        self.date.path = "nothing"
        result = self.date.save_student_checkout('05/29/24', 'A00223332', '14:53', '18:13')
        self.assertEqual(result, (1, 'openpyxl does not support  file format, please check you can open it with Excel first. Supported formats are: .xlsx,.xlsm,.xltx,.xltm'))
        self.date.path = data_path

    def test_get_visits(self):
        self.assertEqual(self.date.get_visits("A00123456"), {"Date":[['Testing', 'Course', 'Section', 'Calc #', 'Time In', 'Time Out']], "04/15/24":[['=FALSE()', '2110', 'N30', '0', '14:52', '16:18']]})
        self.date.path = "nothing"
        self.assertEqual(self.date.get_visits("A00123456"), (1, 'openpyxl does not support  file format, please check you can open it with Excel first. Supported formats are: .xlsx,.xlsm,.xltx,.xltm'))
        self.date.path = data_path
    
    def test_get_max_row(self):
        wb = openpyxl.load_workbook(data_path, read_only=True)
        ws = wb["Date"]
        self.assertEqual(self.date.get_max_row(ws), 16)

    def test_compare(self):
        self.assertEqual(self.date.compare((70, 24), (179, 24)), -1)
        self.assertEqual(self.date.compare((80, 24), (179, 23)), 1)
        self.assertEqual(self.date.compare((180, 24), (180, 24)), 0)

    def test_month(self):
        self.assertEqual(self.date.month(8), 31)


def run_tests():
    '''
    cov = coverage.Coverage()
    cov.start()
    '''
    
    unittest.main(argv=['first-arg-is-ignored'], exit=False)
    
    '''
    cov.stop()
    cov.save()
    cov.html_report(directory="html")
    coverage.CoverageData()
    '''
