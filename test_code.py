import unittest
import dbBot
import mock
from selenium import webdriver
import os

import shutil, tempfile
from time import sleep




class MyTestCase(unittest.TestCase):

    def setUp(self):
        self.driver = webdriver.Chrome('C:\\Users\\yangb\\Desktop\\chromedriver.exe')
        self.driver.set_page_load_timeout(20)
        # Create a temporary directory
        self.test_dir = tempfile.mkdtemp()

    def test_directory_structure(self):
        # create a file in the temporory directory
        self.dbBot.directory_structure('test')
        print(str(os.path.exists('C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBot\\Outputs_test')))
        print(str(os.path.exists(
            'C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBot\\Outputs_test\\Debarment_files_test')))
        print(str(os.path.exists('C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBot\\Outputs_test'
                                 '\\Flagged_authors_files_test')))
        print(str(os.path.exists(
            'C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBot\\Outputs_test\\Completed_file_test')))
        print(str(os.path.exists('C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBot\\Outputs_test'
                                 '\\Screenshots_test')))
    # def test_simepleInput(self):
    #     pageUrl = "http://www.seleniumeasy.com/test/basic-first-form-demo.html"
    #     driver = self.driver
    #     driver.maximize_window()
    #     driver.get(pageUrl)
    #
    #     # Finding "Single input form" input text field by id. And sending keys(entering data) in it.
    #     eleUserMessage = driver.find_element_by_id("user-message")
    #     eleUserMessage.clear()
    #     eleUserMessage.send_keys("Test Python")
    #
    # def test_something(self):
    #     self.assertEqual(dbBot.get_working_filename(), None)
    #
    # def tearDown(self):
    #     self.driver.close()
    #     shutil.rmtree(self.test_dir)


if __name__ == '__main__':
    unittest.main()
