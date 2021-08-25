import unittest

from main import create_ppt_file


class TestCreatePPTFile(unittest.TestCase):
    def test_create_ppt_success(self):
        actual = create_ppt_file("data", "data/nike_black.png")
        expected = "PPT File Generated Successfully"
        self.assertEqual(actual, expected)

    def test_create_ppt_fail(self):
        actual = create_ppt_file("date", "data/nike_black.png")
        expected = "PPT File Generated Successfully"
        self.assertEqual(actual, expected)