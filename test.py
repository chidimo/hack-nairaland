import os
import sys
import unittest
import types

from pathlib import Path
from unittest import mock

from bs4 import BeautifulSoup

from collections import OrderedDict

import hack

TEST_DIRECTORY = Path.joinpath(hack.BASE_DIR, 'test-dir')

# REVISIT
class TestParseCommentBlock(unittest.TestCase):
    pass

    # def test_parse_comment_block(self):
    #     with open(Path.joinpath(TEST_DIRECTORY, 'test_input_comment_parser.html'), 'r+') as rh:
    #         soup = BeautifulSoup(rh.read(), 'html5lib').find('div', class_='narrow')
    #     c = hack.parse_comment_block(soup)

    #     excpected = """
    #     First paragraph

    #     A reply to ...
    #     Another reply, now to
    #     Final reply to poster 3
    #     Final paragraph
    #     ++++++++++++++++

    #     Poster1
    #     Poster 1 first paragraph
    #     Poster 1 second paragraph
    #     Poster 1 3rd paragraph
    #     Poster2
    #     Poster 2 first paragraph
    #     Poster3
    #     Poster 3 first paragraph
    #     """

    #     print("excpected\n", c)
    #     print("parsed_data\n", c)
    #     self.assertEqual(c, excpected)

class TestParsebrTag(unittest.TestCase):
    def test_element_with_only_previous_sibling(self):
        tag_block = """
        This tag has previous sibling text BUT no next sibling.
        <br/>
        """
        tag = BeautifulSoup(tag_block, 'html5lib').find('br')
        self.assertEqual(hack.get_left_right_of_html_br_element(tag)[0], "This tag has previous sibling text BUT no next sibling.")

    def test_element_with_only_next_sibling(self):
        tag_block = """<br/>
        This tag has next sibling text BUT no previous sibling.
        """
        tag = BeautifulSoup(tag_block, 'html5lib').find('br')
        self.assertEqual(hack.get_left_right_of_html_br_element(tag)[1], "This tag has next sibling text BUT no previous sibling.")

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
        self.assertEqual(hack.get_left_right_of_html_br_element(tag), "")
        tag2 = BeautifulSoup(tag_block2, 'html5lib').find('br')
        self.assertEqual(hack.get_left_right_of_html_br_element(tag2), "")

    def test_element_with_both_next_and_previous_siblings(self):
        tag_block = """
        This tag has previous sibling text
        <br/>
        It also has a next sibling text
        """
        tag = BeautifulSoup(tag_block, 'html5lib').find('br')
        self.assertEqual(hack.get_left_right_of_html_br_element(tag), ("This tag has previous sibling text", "It also has a next sibling text"))

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
    def test_scrap_comments_for_range_of_post_pages(self, mocked_ripper):
        """Test length and type of object returned"""
        obj = hack.PostCollector(self.base_url)
        page_range = obj.scrap_comments_for_range_of_post_pages(0, 3, False)
        self.assertIsInstance(page_range, types.GeneratorType)
        self.assertEqual(len(list(page_range)), 4)

        # Set maximum page to 3
        obj.max_page = mock.MagicMock(return_value=3)
        page_range = obj.scrap_comments_for_range_of_post_pages(0, 0, True)
        self.assertEqual(len(list(page_range)), 3)

if __name__ == '__main__':
    print("Testing")
    unittest.main()
