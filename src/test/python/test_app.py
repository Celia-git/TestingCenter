import unittest

from .__init__ import *


class AppTester(unittest.TestCase):

    def __init__(self, methodName: str = "runTest") -> None:
        super().__init__(methodName)
        self.app = App()
        self.container_frame = self.app.container_frame
        self.default_frame = self.container_frame.load_default_frame()
        self.student_frame = self.container_frame.load_student_frame()
        self.date_frame = self.container_frame.load_date_frame()

    async def _start_app(self):
        self.app.mainloop()

    def setUp(self):
        self._start_app()

    def test_startup(self):
        title = self.app.winfo_toplevel().title()
        expected = "Student Scanner"
        self.assertEqual(title, expected)

    def test_container_frame(self):
        self.assertIsInstance(self.container_frame, tk.Frame)

    def test_load_default_frame(self):
        default_frame = self.app.container_frame.load_default_frame()
        self.assertIsInstance(default_frame, tk.Frame)

    def test_load_student_frame(self):
        student_frame = self.app.container_frame.load_student_frame()
        self.assertIsInstance(student_frame, tk.Frame)

    def test_date_frame(self):
        self.date_frame = self.app.container_frame.load_date_frame()
        self.assertIsInstance(self.date_frame, tk.Frame)

    def test_swap_default(self):
        self.container_frame.swap(0)
        self.assertEqual(self.container_frame.top_frame, self.default_frame)

    def test_swap_student(self):
        self.container_frame.swap(1)
        self.assertEqual(self.container_frame.top_frame, self.student_frame)
    
    def test_swap_date(self):
        self.container_frame.swap(2)
        self.assertEqual(self.container_frame.top_frame, self.date_frame)

    def test_is_valid_Anumber(self):
        self.container_frame.a_num.set("A00123789")
        self.assertTrue(self.container_frame.is_valid_Anumber())
        self.container_frame.a_num.set("A0012345678")
        self.assertFalse(self.container_frame.is_valid_Anumber())
        self.container_frame.a_num.set("A0012")
        self.assertFalse(self.container_frame.is_valid_Anumber())
        self.container_frame.a_num.set("123456789")
        self.assertFalse(self.container_frame.is_valid_Anumber())
    
    def test_submit_valid(self):
        self.container_frame.a_num.set("A00123789")
        self.default_frame.submit_but.invoke()
        self.assertEqual(self.container_frame.top_frame, self.student_frame)

    def test_submit_invalid(self):
        self.container_frame.a_num.set("A0012")
        self.default_frame.submit_but.invoke()
        self.assertEqual(self.container_frame.top_frame, self.default_frame)
        self.assertEqual(self.default_frame.error_label.cget("text"), "Error: A0012: not a valid A Number\nEnter a student ID which starts with A00 and ends with six digits")
    
    def test_view(self):
        self.default_frame.view_but.invoke()
        self.assertEqual(self.container_frame.top_frame, self.date_frame)

    def test_load_student(self):
        self.student_frame.discard(True)
        self.container_frame.a_num.set("A00444555")
        self.container_frame.load_student_frame()
        self.assertIsInstance(self.student_frame.student, Student)
        self.assertIsInstance(self.student_frame.date_log, Date)
        self.assertEqual(self.container_frame.a_num.get(), "A00444555")
        self.assertEqual(self.student_frame.name.get(), "Morgan Curry")

    # Test load_student-> course_values
    def test_write_courses(self):
        self.student_frame.discard(True)
        self.container_frame.a_num.set("A00444555")
        self.container_frame.load_student_frame()
        self.assertEqual(self.student_frame.cor_value["values"], ('1910', '---------------------', '1030', '1130', '1710', '1720', '1910', '1920', '2010', '2110'))
        self.assertEqual(self.student_frame.sec_value["values"], ('N30', 'N35'))

    def test_write_visits(self):
        self.student_frame.discard(True)
        self.student_frame.load_student("A00444555")
        label_text = ['Date', 'Testing', 'Course', 'Section', 'Calc #', 'Time In', 'Time Out', '03/08/24', 'TRUE', '1910', 'N35\n', '1', '13:59', '17:00', '04/10/24', 'TRUE', '1910', 'N35', '12', '14:51', '15:56', '04/17/24', 'TRUE', '1910', 'N35', '0', '13:17', '17:13']
        self.assertEqual([x.cget("text") for x in self.student_frame.display_frame.winfo_children()], label_text)

    def test_save_student_values(self):
        pass

    def test_discard_student_entries(self):
        pass


    def tearDown(self) -> None:
        self.app.destroy()


def run_tests():
    unittest.main(argv=['first-arg-is-ignored'], exit=False)