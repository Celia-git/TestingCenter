import unittest

from .__init__ import *


class AppTester(unittest.TestCase):

    def __init__(self, methodName: str = "runTest") -> None:
        super().__init__(methodName)
        self.app = App()
    
    def test_load_default_frame(self):
        self.assertIsInstance(self.app.container_frame.load_default_frame(), tk.Frame)


        
def run_tests():
    unittest.main(argv=['first-arg-is-ignored'], exit=False)