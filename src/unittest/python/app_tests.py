import unittest

from __init__ import *


class AppTester(unittest.TestCase):

    def __init__(self, methodName: str = "runTest") -> None:
        super().__init__(methodName)
        self.app = App()
    
        
        
if __name__ == "__main__":
    unittest.main()