import unittest

from MyClass.Data import Data

class TestData(unittest.TestCase):
    
    def Test_GetFileName(self):

        data = Data("Data.xlsx")
        result = data.GetFileName(1)
        self.assertEqual(result, "Pippo")
        # self.assertEqual(data.GetFileName(2), "Pluto")
        # self.assertEqual(data.GetFileName(3), "Paperino")
