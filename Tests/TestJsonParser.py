import unittest
import FileParser.JsonParser as jp

class MyTestCase(unittest.TestCase):
    def test_scanAllFiles(self):
        jp.scanAllFiles('G:\养牛记录\联曼')
        self.assertEqual(True, True)  # add assertion here


if __name__ == '__main__':
    unittest.main()
