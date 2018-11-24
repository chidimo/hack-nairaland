import os
import sys
import unittest
import types
from unittest import mock

from bs4 import BeautifulSoup

from collections import OrderedDict

import hack

# REVISIT
class TestParseCommentBlock(unittest.TestCase):
    def test_comment_block_parsing(self):
        test_ouput_directory = os.path.join(hack.BASE_DIR, 'test-dir')

        with open('test-dir/test_input_comment_parser.html', 'r+') as rh:
            soup = BeautifulSoup(rh.read(), 'html5lib').find('div', class_='narrow')
        with open('test-dir/test_comment_parser_output.txt', 'r+') as r:
            excpected = r.read()
        parsed_data = hack.parse_comment_block(soup)

        print(excpected)
        print()
        print(parsed_data)
        self.assertEqual(parsed_data, excpected)

class TestParsebrTag(unittest.TestCase):
    def test_tag_with_only_previous_sibling(self):
        tag_block = """
        This tag has previous sibling text BUT no next sibling.
        <br/>
        """
        tag = BeautifulSoup(tag_block, 'html5lib').find('br')
        self.assertEqual(hack.parse_html_br_tag_content(tag)[0], "This tag has previous sibling text BUT no next sibling.")

    def test_tag_with_only_next_sibling(self):
        tag_block = """<br/>
        This tag has next sibling text BUT no previous sibling.
        """
        tag = BeautifulSoup(tag_block, 'html5lib').find('br')
        self.assertEqual(hack.parse_html_br_tag_content(tag)[1], "This tag has next sibling text BUT no previous sibling.")

    def tag_with_no_previous_or_next_sibling(self):
        tag_block = """
        <br/>
        <br/>
        """
        tag_block2 = """
        <br/>
        <br/>
        <br/>
        """
        tag = BeautifulSoup(tag_block, 'html5lib').find('br')
        self.assertEqual(hack.parse_html_br_tag_content(tag), "")
        tag2 = BeautifulSoup(tag_block2, 'html5lib').find('br')
        self.assertEqual(hack.parse_html_br_tag_content(tag2), "")

    def test_tag_with_both_next_and_previous_siblings(self):
        tag_block = """
        This tag has previous sibling text
        <br/>
        It also has a next sibling text
        """
        tag = BeautifulSoup(tag_block, 'html5lib').find('br')
        self.assertEqual(hack.parse_html_br_tag_content(tag), ("This tag has previous sibling text", "It also has a next sibling text"))

class TestPostCollector(unittest.TestCase):
    def setUp(self):
        self.base_url = "some/base/url"

    @mock.patch('hack.os')
    def test_init(self, mock_os):
        # Save directory exists; 'mkdir' NOT called
        mock_os.path.join.return_value = 'joined-path'
        mock_os.path.exists.return_value = True
        obj = hack.PostCollector(self.base_url)
        self.assertEqual(obj.title, 'url')
        self.assertFalse(mock_os.mkdir.called, "Directory creation method called")

        # Save directory doesn't exists; 'mkdir' called
        mock_os.path.join.return_value = 'joined-path'
        mock_os.path.exists.return_value  = False
        hack.PostCollector(self.base_url)
        mock_os.mkdir.assert_called_with('joined-path')

    @mock.patch('hack.requests')
    def test_url_checker(self, mocked_requests):
        check = hack.PostCollector(self.base_url)._check_if_url_exists_and_is_valid('whatever-url')
        # Assert that requests head method was called during function execution
        mocked_requests.head.assert_called_with("whatever-url")
        self.assertIsInstance(check, bool)

    @mock.patch('hack.requests')
    def test_max_page(self, mocked_requests):
        obj = hack.PostCollector(self.base_url)
        obj._check_if_url_exists_and_is_valid = mock.MagicMock(return_value=False)
        self.assertEqual(obj.max_page(), 0)

    @mock.patch('hack.Ripper')
    def test_scrap_comment_for_single_page(self, mocked_ripper):
        """Test return object type"""
        obj = hack.PostCollector(self.base_url)
        single_page = obj._scrap_comment_for_single_page("page-url")
        mocked_ripper.assert_called()
        self.assertIsInstance(single_page, OrderedDict)

    @mock.patch('hack.Ripper')
    def test_scrap_comments_for_range_of_pages(self, mocked_ripper):
        """Test length and type of object returned"""
        obj = hack.PostCollector(self.base_url)
        page_range = obj.scrap_comments_for_range_of_pages(0, 3, False)
        self.assertIsInstance(page_range, types.GeneratorType)
        self.assertEqual(len(list(page_range)), 4)

        # Set maximum page to 3
        obj.max_page = mock.MagicMock(return_value=3)
        page_range = obj.scrap_comments_for_range_of_pages(0, 0, True)
        self.assertEqual(len(list(page_range)), 3)

if __name__ == '__main__':
    print("Testing")
    unittest.main()
