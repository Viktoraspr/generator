"""
This file is for testing classes and their methods
"""
import os
import unittest

from src.generator import Generator


class TestFileManagement(unittest.TestCase):
    """
    Testing FileManagement and DBManagement classes
    """
    excel_name = r'C:\Pycharm projects\generator\tests\generator.xlsx'
    generator = Generator(excel_path=excel_name)

    def test_create_table_with_selected_columns(self):
        """
        Checks if json file was created.
        If exits, file is deleting
        """
        json_path = self.generator.run()
        assert os.path.exists(json_path)

        os.remove(json_path)
