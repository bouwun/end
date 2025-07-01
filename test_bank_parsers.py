import unittest
from bank_parsers import ICBCParser

class TestICBCParser(unittest.TestCase):
    def setUp(self):
        self.parser = ICBCParser()
        
    def test_parse_date(self):
        self.assertEqual(self.parser.parse_date("2023-01-01"), "2023-01-01")
        self.assertEqual(self.parser.parse_date("2023/01/01"), "2023-01-01")

if __name__ == "__main__":
    unittest.main()