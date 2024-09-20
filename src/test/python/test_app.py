import unittest

from .__init__ import *


class AppTester(unittest.TestCase):

    async def _start_app(self):
        self.app.mainloop()

    def setUp(self):
        self.app = App()
        self._start_app()

    def tearDown(self) -> None:
        self.app.destroy()

    def test_startup(self):
        title = self.app.winfo_toplevel().title()
        expected = "Student Scanner"
        self.assertEqual(title, expected)


def run_tests():
    unittest.main(argv=['first-arg-is-ignored'], exit=False)