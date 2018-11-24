import re
import os
import sys
import codecs
import logging

from operator import itemgetter
from itertools import filterfalse
from collections import OrderedDict, namedtuple, Counter

import bs4
import docx
import requests
import requests.exceptions as rqe
import openpyxl as OP

from pywebber import Ripper

logging.disable(logging.CRITICAL)

USERHOME = os.path.abspath(os.path.expanduser('~'))
DESKTOP = os.path.abspath(USERHOME + '/Desktop/')
BASE_DIR = STATUS_DIR = os.path.join(DESKTOP ,"hack-nairaland")

if not os.path.exists(BASE_DIR):
    os.mkdir(BASE_DIR)

class Error(Exception):
    pass

class NonExistentNairalandUser(Error):
    pass

def new_logger(log_file_name):
    FORMATTER = logging.Formatter("%(asctime)s:%(funcName)s:%(levelname)s\n%(message)s")
    # console_logger = logging.StreamHandler(sys.stdout)
    file_logger = logging.FileHandler(log_file_name)
    file_logger.setFormatter(FORMATTER)

    logger = logging.getLogger(log_file_name)
    logger.setLevel(logging.DEBUG)
    logger.addHandler(file_logger)
    logger.propagate = False
    return logger

PARSE_BR_TAG_LOGGER = new_logger('log_html_br_tag.log')
PARSE_COMMENT_BLOCK_LOGGER = new_logger('log_parse_comment_block.log')
FORMAT_COMMENTS_LOGGER = new_logger('log_format_comments.log')

def check_if_url_exists_and_is_valid(url):
    r = requests.head(url)
    return r.status_code == 200

def unique_everseen(iterable, key=None):
    """List unique elements, preserving order. Remember all elements ever seen.
    source: https://docs.python.org/3/library/itertools.html#itertools-recipes"""
    seen = set()
    seen_add = seen.add
    if key is None:
        for element in filterfalse(seen.__contains__, iterable):
            seen_add(element)
            yield element
    else:
        for element in iterable:
            k = key(element)
            if k not in seen:
                seen_add(k)
                yield element

def dict_to_string(dictionary):
    """Flatten a dictionary into a single string"""
    if isinstance(dictionary, dict):
        return '\n'.join([' says\n'.join([key, value]) for key, value in dictionary.items()])
    return None

def parse_html_br_tag_content(break_tag):
    """Get content of the next and previous sibling of a <br/> tag

    Parameters
    -----------
    BeautifulSoup
        BeautifulSoup object of <br/> tag

    Returns
    -------
    tuple
        A tuple of (previous_sibling_text, next_sibling_text)

    Notes
    ------
    1. For each <br/> tag we run the .string method on its next and previous siblings
    2. If a proper string is encountered, it is returned.
    3. If <br/> is encountered, None is returned.
    4. If any other string which result in an error is encountered, None is returned
    """
    p_sibling = break_tag.previous_sibling
    n_sibling = break_tag.next_sibling

    PARSE_BR_TAG_LOGGER.debug("previous sibling\n{}".format(p_sibling))
    PARSE_BR_TAG_LOGGER.debug("next sibling\n{}".format(n_sibling))

    return_value = [None, None]
    try:
        return_value[1] = n_sibling.string.strip().strip("\n:")
    except AttributeError:
        pass
    try:
        return_value[0] = p_sibling.string.strip().strip("\n:")
    except AttributeError:
        pass
    # return ("*{}*, *{}*".format(return_value[0], return_value[1]))
    return tuple(return_value)

def join_tuples(list_of_tuples):
    """Join a list of tuples into a list eliminating None and
    duplicates.

    Parameters
    -----------
    list
        A list of tuples whose elements are strings.

    Returns
    --------
    list
        A list of unique strings in the input list

    Notes
    ------
    unique_everseen removes duplicates.
    This is needed in cases where a next_sibling and a
    previous_sibling point to the same string.
    """
#     remove_nones = [[filter(lambda x: x is not None, each)] for each in list_of_tuples]
    
#     # a more explicit way
#     remove_nones2 = []
#     for each_tuple in list_of_tuples:
#         l = []
#         for each_string in each_tuple:
#             if each_string is not None:
#                 l.append(each_string)
#         remove_nones2.append(l)
#     print(remove_nones2)

    # use list comprehension for speed
    remove_nones = [
        [each_string for each_string in each_tuple if each_string is not None] for each_tuple in list_of_tuples
    ]
    phrase_collection = [phrase.strip().strip("\n:") for each in remove_nones for phrase in each]
    return "\n".join(unique_everseen(phrase_collection))

def format_comments(bs4_comment_block_object):
    """Format a comment block into proper paragraphs

    Parameters
    ------------
    BeautifulSoup
        BeautifulSoup object of comment block

    Returns
    --------
    str
        A properly paragraphed string
    """

    FORMAT_COMMENTS_LOGGER.debug(bs4_comment_block_object.prettify())

    comment = []
    break_tags = bs4_comment_block_object.find_all('br')

    if break_tags == []:
        return bs4_comment_block_object.text
    for each in break_tags:
        content = parse_html_br_tag_content(each)
        comment.append(content)

    return_string = join_tuples(comment)
    return return_string

def parse_comment_block(bs4_comment_block_object):
    """Return quoted string.
    
    Parameters
    -----------
    BeautifulSoup
        BeautifulSoup object of a quoted string block

    Returns
    --------
    quoted : OrderedDict()
        quoted string content
    bs4_comment_block_object : BeautifulSoup
        Input comment block stripped of all <b> tags
        
    Notes
    ------
    Every comment block must be parsed with this function.
    This function also has a side effect of producing a properly formatted html of all comments it encounters.
    """
    
    PARSE_COMMENT_BLOCK_LOGGER.debug(bs4_comment_block_object.prettify())

    save_dir = os.path.join(BASE_DIR, "comment_block")
    if os.path.exists(save_dir) is False:
        os.mkdir(save_dir)

    # Side effect
    save = os.path.join(save_dir, "all_page_comments.html")
    with codecs.open(save, 'a+', encoding='utf-8') as f:
        f.write(bs4_comment_block_object.prettify())
        f.write("<div class='dropdown-divider'></div>")

    collected_quotes = OrderedDict()
    return_val = namedtuple('ParsedComment', ['focus_user_comment', 'quotes_ordered_dict'])
    blockquotes = bs4_comment_block_object.find_all('blockquote')

    if blockquotes == []:
        pass
    else:
        for blockquote in blockquotes:
            try:
                commenter = blockquote.find('b').text
            except AttributeError:
                commenter = 'Anonymous'

            collected_quotes[commenter] = format_comments(blockquote).strip().strip("\n:")
            blockquote.decompose() # remove the block from the tree

    return_val.focus_user_comment = format_comments(bs4_comment_block_object).strip().strip("\n:")
    return_val.quotes_ordered_dict = collected_quotes
    return return_val

def sort_dictionary_by_value(dictionary_to_sort):
    """
    Return list of dictionary keys where the items are sorted on the values in descending order.
    
    e.g sort_dictionary_by_value({5:'goat', 10:'cat', 1:'dog'}) returns [5, 1, 10] since the values
    sort to ['goat', 'dog', 'cat']
    """
    if dictionary_to_sort is None:
        return
    ordered_dictionary = sorted(dictionary_to_sort.items(), key=itemgetter(1), reverse=True)
    sorted_dictionary_list = [i[0] for i in ordered_dictionary]
    return sorted_dictionary_list

class PostCollector:
    """
    Scrap a nairaland post

    Parameters
    ----------
    str
        Post url
    """

    def __str__(self):
        return "PostCollector: {}".format(self.base_url)

    def __init__(self, base_url, refresh=True):
        self.save_path = os.path.join(BASE_DIR, 'page_rips_post')
        self.base_url = base_url # Page (0) of the post
        self.refresh = refresh
        self.title = self.base_url.split('/')[-1]
        if not os.path.exists(self.save_path):
            os.mkdir(self.save_path)

    def max_page(self):
        """Returns the maximum number of pages of comments, starting from a zero index."""
        stop = 0
        while True:
            print("Whiling in a loop")
            if self._check_if_url_exists_and_is_valid("{}/{}".format(self.base_url, stop)): # check if next url exists
                stop += 1
            else:
                break
        return stop

    @staticmethod
    def _check_if_url_exists_and_is_valid(url): # ConnectionError happens here
        r = requests.head(url)
        return r.status_code == 200

    def _scrap_comment_for_single_page(self, page_url):
        """Return comments and commenters on a single post page

        Returns
        --------
        OrderedDict
            Dictionary of {commenter : comments}
        """

        # User posts are contained in a table with summary='posts' attribute.
        # Each commenter name is contained inside a <tr>
        # Each comment is contained in <tr> just below the name of the commenter
        soup = Ripper(page_url, parser='html5lib', save_path=self.save_path, refresh=self.refresh).soup

        # Handle supposed anomaly by decomposing all such occurrences from the tree
        for each in soup.find_all('td', class_="l pu pd"):
            each.parent.decompose()
        rows = soup.find('table', summary='posts').find_all('tr')

        return_val = OrderedDict()
        for i in range(0, len(rows), 2):

            topic_classes = ['bold l pu', 'bold l pu nocopy'] # topic div should be either of these classes
            for class_ in topic_classes:
                try:
                    moniker = rows[i].find('td', class_=class_).find('a', href=True, class_=True).text.strip()
                    break
                except AttributeError:
                    pass
                moniker = "Nobody" # set to nobody after exhausting all options. We cannot use finally in this case

            comment_classes = ['l w pd', 'l w pd nocopy'] # comment div should be either of these classes            
            for class_ in comment_classes:
                try:
                    comment_block = rows[i+1].find('td', id=True, class_=class_).find('div', class_='narrow')
                    break
                except AttributeError:
                    pass
            parsed_block = parse_comment_block(comment_block)

            # If a moniker already exists (i.e. a user has already commented), append an integer to the
            # present one to differentiate both.
            if moniker not in return_val:
                return_val[moniker] = parsed_block
            else:
                moniker = "{}**{}".format(moniker, i)
                return_val[moniker] = parsed_block
        return return_val

    def scrap_comments_for_range_of_pages(self, start=0, stop=1, __all=False):
        """Get contents for a range of pages from start to stop"""
        if __all == True: # since we're starting from a zero index, we have to subtract 1 from self.max_page()
            stop = self.max_page() - 1
        while start <= stop:
            next_url = "{}/{}".format(self.base_url, start)
            next_page = self._scrap_comment_for_single_page(next_url)
            yield next_page
            start += 1

    def all_commenters(self):
        """Return list of all commenters on a post"""
        # Remember we user ** to separate a moniker and the number of times it is appearing on a post
        return sorted([key.split("**")[0] for each in list(self.scrap_comments_for_range_of_pages(stop=self.max_page())) for key, value in each.items()])

    def unique_commenters(self):
        """Return list of unique commenters on a post"""
        return sorted(set(self.all_commenters()))

    def commenters_activity_summary(self):
        """Return count of number of times a user commented on a post"""
        x = Counter(self.all_commenters())
        print(x)
        print("Sorted dict")
        print(sort_dictionary_by_value(x))
        return 

class UserCommentHistory:
    """
    Grab a user's comment history

    Parameters
    ------------
    user_name : str
        User's name to crawl. Default is 'seun'
    """

    def __str__(self):
        return "UserCommentHistory: {}".format(self.user_post_page)

    def __init__(self, nairaland_moniker, refresh=True):
        self.refresh = refresh
        BASE_URL = 'https://www.nairaland.com'
        self.save_path = os.path.join(BASE_DIR, 'page_rips_user')
        if not os.path.exists(self.save_path):
            os.mkdir(self.save_path)
            
        p = '{}/{}'.format(BASE_URL, nairaland_moniker.lower())
        if self._check_if_url_exists_and_is_valid(p):
            self.user_profile_page = p
            self.user_post_page = '{}/{}/posts'.format(BASE_URL, nairaland_moniker.lower())
        else:
            raise NonExistentNairalandUser("This user does not exist on nairaland.")

    @staticmethod
    def _check_if_url_exists_and_is_valid(url):
        r = requests.head(url)
        return r.status_code == 200
    
    def user_profile(self):
        """Returns a dictionary of the user's profile"""
        pass

    def max_pages(self):
        """Return number of pages of comment for user"""
        soup = Ripper(self.user_post_page, save_path=self.save_path, refresh=True).soup
        pattern = r"\<b\>\s*(\d+)\s*\<\/b\>" # pattern to search for number of pages of comments
        try:
            return int(re.search(pattern, str(soup)).group(1))
        except AttributeError:
            print("Could not find max page")
            pass

    def _scrap_comment_for_single_page(self, page_url):
        """Return comments and commenters on a single post page

        Returns
        --------
        OrderedDict
            Dictionary of {section : namedtuple}
        """

        # User comments are contained in a table with neither summary nor id attribute.
        # Then follows the rows containing the section, topic, and username and the comment itself just below it
        soup = Ripper(page_url, parser='html5lib', save_path=self.save_path, refresh=self.refresh).soup

        for each in soup.find_all('td', class_="l pu pd"):
            each.parent.decompose() # remove these trees as they are unneeded
        rows = soup.find('table', id=False, summary=False).find_all('tr')

        return_val = OrderedDict()
        for i in range(0, len(rows), 2): # go to every second row
            
            topic_classes = ['bold l pu', 'bold l pu nocopy']
            for class_ in topic_classes:
                try:
                    section_topic = rows[i].find('td', class_=class_).find_all('a', href=True, class_=False)
                except AttributeError:
                    pass

            comment_classes = ['l w pd', 'l w pd nocopy']
            for class_ in comment_classes:
                try:
                    comment_block = rows[i+1].find('td', class_=class_).find('div', class_='narrow')
                except AttributeError:
                    pass

            section = section_topic[0].text.strip()
            topic = section_topic[1].text.lstrip("Re:").strip()

            parsed_block = parse_comment_block(comment_block)

            Comm = namedtuple('Comment', ['topic', 'parsed_comment'])
            Comm.topic = topic
            Comm.parsed_comment = parsed_block

            if section not in return_val:
                return_val[section] = Comm
            else:
                section = "{}**{}".format(section, i)
                return_val[section] = Comm
        return return_val

    def scrap_comments_for_range_of_pages(self, start=0, stop=0, _maximum_pages=False):
        """Get contents for a range of pages from start to stop """
        if _maximum_pages:
            stop = self.max_pages() - 1
        while start <= stop:
            next_url = "{}/{}".format(self.user_post_page, start)
            next_page = self._scrap_comment_for_single_page(next_url)
            yield next_page
            start += 1

class TopicCollector:
    """
    Collect topics from a section of nairaland

    Parameters
    -----------
    section : str
        Default section is politics

    Notes
    ------
    The methods in this class are only applicable to section urls
    """

    def __str__(self):
        return "TopicCollector: {}".format(self.base_url)

    def __init__(self, section='politics'):
        self.save_path = os.path.join(BASE_DIR, 'page_rips_section')
        self.section = section
        self.base_url = 'http://www.nairaland.com/{}'.format(self.section)
        if os.path.exists(self.save_path) is False:
            os.mkdir(self.save_path)

    def max_pages(self):
        """Return number of pages in this section

        Returns
        --------
        int
            The number of pages in this section
        """
        soup = Ripper(self.base_url, parser='html5lib', save_path=self.save_path, refresh=True).soup
        number = re.search(r"\(of\s*(\d+)\s*pages\)", soup.text).group(1)
        return int(number)

    def _scrap_topics_for_a_single_page(self, page_url, refresh=True):
        """
        Yield all topics on a page

        Yields
        -------
        namedtuple
            collection of 'poster', 'title', 'url', 'number of comments'
        """
        soup = Ripper(page_url, parser='html5lib', save_path=self.save_path, refresh=True).soup
        post_table = soup.find('table', id=False, summary=False)

        for td in post_table.find_all('td', id=True):
            Post = namedtuple('Post', ['poster', 'title', 'url', 'comments'])

            title_component = td.find('b').find('a', href=True)
            Post.title = title_component.text.strip()
            Post.url = 'http://www.nairaland.com' + title_component.get('href').strip()

            # there is a maximum of 7 <b> tags
            meta_component = td.find('span', class_='s').find_all('b')

            Post.poster = meta_component[0].text.strip()
            Post.comments = int(meta_component[1].text) # count includes the post itself
            yield Post

    def scrap_topics_for_range_of_pages(self, start=0, stop=0, _maximum_pages=False):
        """Yield all topics between 'start' and 'end' for a section

        Parameters
        -----------
        int
            Start and end values of section

        Yields
        -------
        tuple
            same yields as for titles()
        """
        if _maximum_pages:
            stop = self.max_pages() - 1
        while start <= stop:
            next_url = '{}/{}'.format(self.base_url, start)
            yield self._scrap_topics_for_a_single_page(next_url)
            start += 1

def export_user_comments_to_html(username=None, max_page=5):
    """Export all of a user's comments data to a html file

    Parameters
    -----------
    str
        Username
    int
        Maximum page count for user's comments (Default is 5 pages of comments)
        loop breaks if we exceed actual count
    """
    
    if not username:
        print("No username provided. Ending")
        return
    else:
        print("Now hacking nairaland. Please wait a few minutes.")
        
    destination_file = os.path.join(BASE_DIR, "comments_{}_{}_pages.html".format(username.lower(), max_page))
    if os.path.exists(destination_file):
        os.remove(destination_file)
    with open(destination_file, 'a+', encoding='utf-8') as f:
        
        # html scaffold
        f.write("<html xmlns='http://www.w3.org/1999/xhtml'>\n")
        f.write("\t<head>\n")

        
        # resources
        f.write("\t\t<link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootswatch/4.1.3/superhero/bootstrap.min.css'>\n")
        f.write("\t\t<meta name='viewport' content='width=device-width, initial-scale=1, shrink-to-fit=no'>\n")
        f.write("\t\t<script src='https://stackpath.bootstrapcdn.com/bootstrap/4.1.1/js/bootstrap.min.js' integrity='sha384-smHYKdLADwkXOn1EmN1qk/HfnUcbVRZyYmZ4qpPea6sjB/pTJ0euyQp0Mk8ck+5T' crossorigin='anonymous' async></script>\n")
        f.write("<script defer src='https://use.fontawesome.com/releases/v5.0.6/js/all.js' async></script>")
        f.write("<script src='https://code.jquery.com/jquery-3.3.1.min.js' integrity='sha256-FgpCb/KJQlLNfOu91ta32o/NMZxltwRo8QtmkMRdAu8=' crossorigin='anonymous'></script>")
        # end resources

        f.write("\t\t<title>Comment history for {} - Hack Nairaland</title>\n".format(username.lower()))
        f.write("\t</head>\n")
        f.write("\t<body style='padding-top:7rem;margin-bottom:5rem;'>\n")
        # navbar
        f.write("\t\t<nav class='navbar navbar-expand-lg navbar-dark bg-primary fixed-top' id='topNav'>\n")
        f.write("\t\t<a class='navbar-brand' style='font-size:36px;'>Hack Nairaland</a>\n")
        f.write("\t\t<button type='button' class='navbar-toggler my-toggler' data-toggle='collapse' data-target='.navcontent'>\n")
        f.write("\t\t<span class='sr-only'>Toggle navigation</span>\n")
        f.write("\t\t<span class='navbar-toggler-icon'></span>\n")
        f.write("\t\t</button>\n")
        f.write("\t\t<div class='collapse navbar-collapse navcontent'>\n")
        f.write("\t\t<ul class='nav navbar-nav lefthand-navigation'>\n")
        f.write("\t\t<li class='nav-item'><a class='nav-link' href='#' title='Home'>Home</a></li>\n")
        f.write("\t\t</ul>\n")
        f.write("\t\t</div>\n")
        f.write("\t\t</nav>\n")
        # end navbar
        f.write("\t\t<div class='container'>\n")
        f.write("<h1>Nairaland comment history for <a href='https://nairaland.com/{0}/posts' target='_blank'>{0}</a></h1>\n".format(username))
        # breadcrumb
        f.write("\t\t<nav aria-label='breadcrumb'>\n")
        f.write("\t\t<ol class='breadcrumb'>\n")
        f.write("\t\t<li class='breadcrumb-item'><a href='#'>Home</a></li>\n")
        f.write("\t\t<li class='breadcrumb-item'>The first {} pages</li>\n".format(max_page))
        f.write("\t\t</ol>\n")
        f.write("\t\t</nav>\n")
        # end breadcrumb

        i = 1
        user = UserCommentHistory(username)
        for page in list(user.scrap_comments_for_range_of_pages(stop=max_page)):

            f.write("\t\t\t<div id='js-scroll-target{}'>\n".format(i)) # div for targeting scroll
            f.write("\t\t\t<h2><a href='#js-scroll-target{0}' class='smooth-scroll'>Page {1}</a></h2>".format(i+1, i))
            i += 1

            for section, topic_plus_comment in page.items():
                f.write("\t\t\t<h3>Section: {}</h3>\n".format(section.split('**')[0])) # remove the ** separating section and index
                f.write('\t\t\t<h4>Subject: {}</h4>'.format(topic_plus_comment.topic))
                
                parsed_comment = topic_plus_comment.parsed_comment
                f.write("\t\t\t<p class='text-success'>{}</p>\n".format(parsed_comment.focus_user_comment))
                quotes = parsed_comment.quotes_ordered_dict
                
                for username, comment in quotes.items():
                    f.write("\t\t\t\t<h4 class='text-info'>{}</h4>\n".format(username))
                    f.write("\t\t\t\t<p class='text-primary'><em>{}</em></p>\n".format(comment))
                f.write("\t\t\t<div class='dropdown-divider' style='border:1px solid white;'></div>\n")
            f.write("\t\t\t</div>\n") # finish scroll div

        # continue up page structure
        f.write("\t\t\t<p class='float-right'><a href='#topNav' class='smooth-scroll'>Back to top</a></p>\n") # back to top
        f.write("\t\t</div>\n")
        # jquery smooth scroll
        f.write("<script>\n")
        f.write("$(document).ready(function(){\n")
        f.write("\t $('.smooth-scroll').on('click', function(event) {\n")
        f.write("\t\tif (this.hash !== '') {\n")
        f.write("\t\t\tevent.preventDefault();\n")
        f.write("\t\t\tvar hash = this.hash;\n")
        f.write("\t\t\t$('html, body').animate({\n")
        f.write("\t\t\t scrollTop: $(hash).offset().top\n")
        f.write("\t\t\t}, 800, function(){\n")
        f.write("\t\t\t window.location.hash = hash;\n")
        f.write("\t\t\t});\n")
        f.write("\t\t}\n")
        f.write("\t});\n")
        f.write("});\n")
        f.write("</script>\n")
        # end jquery smooth scroll
        f.write("\t</body>\n")
        f.write("</html>")
    print("Done hacking")
    os.startfile(destination_file)

def export_topics_to_html(section='romance', start_page=0, stop_page=3):
    """
    Writes all topics between start and end of a section to excel.
    Same output as titles_links_metadata() but written to a excel file
    """
    
    print("Now hacking nairaland. Please wait a few minutes.")
        
    destination_file = os.path.join(BASE_DIR, "{}_page_{}_{}_pages.html".format(section, start_page, stop_page))
    if os.path.exists(destination_file):
        os.remove(destination_file)
    with open(destination_file, 'a+', encoding='utf-8') as f:
        
        # html scaffold
        f.write("<html xmlns='http://www.w3.org/1999/xhtml'>\n")
        f.write("\t<head>\n")
        
        # resources
        f.write("\t\t<link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootswatch/4.1.3/superhero/bootstrap.min.css'>\n")
        f.write("\t\t<meta name='viewport' content='width=device-width, initial-scale=1, shrink-to-fit=no'>\n")
        f.write("\t\t<script src='https://stackpath.bootstrapcdn.com/bootstrap/4.1.1/js/bootstrap.min.js' integrity='sha384-smHYKdLADwkXOn1EmN1qk/HfnUcbVRZyYmZ4qpPea6sjB/pTJ0euyQp0Mk8ck+5T' crossorigin='anonymous' async></script>\n")
        f.write("<script defer src='https://use.fontawesome.com/releases/v5.0.6/js/all.js' async></script>")
        f.write("<script src='https://code.jquery.com/jquery-3.3.1.min.js' integrity='sha256-FgpCb/KJQlLNfOu91ta32o/NMZxltwRo8QtmkMRdAu8=' crossorigin='anonymous'></script>")
        # end resources

        f.write("\t\t<title>Topics filed under {} - Hack Nairaland</title>\n".format(section))
        f.write("\t</head>\n")
        f.write("\t<body style='padding-top:7rem;margin-bottom:5rem;'>\n")
        # navbar
        f.write("\t\t<nav class='navbar navbar-expand-lg navbar-dark bg-primary fixed-top' id='topNav'>\n")
        f.write("\t\t<a class='navbar-brand' style='font-size:36px;'>Hack Nairaland</a>\n")
        f.write("\t\t<button type='button' class='navbar-toggler my-toggler' data-toggle='collapse' data-target='.navcontent'>\n")
        f.write("\t\t<span class='sr-only'>Toggle navigation</span>\n")
        f.write("\t\t<span class='navbar-toggler-icon'></span>\n")
        f.write("\t\t</button>\n")
        f.write("\t\t<div class='collapse navbar-collapse navcontent'>\n")
        f.write("\t\t<ul class='nav navbar-nav lefthand-navigation'>\n")
        f.write("\t\t<li class='nav-item'><a class='nav-link' href='#' title='Home'>Home</a></li>\n")
        f.write("\t\t</ul>\n")
        f.write("\t\t</div>\n")
        f.write("\t\t</nav>\n")
        # end navbar
        f.write("\t\t<div class='container'>\n")
        f.write("<h1>Topics filed under <a href='https://nairaland.com/{0}' target='_blank'>{0}</a></h1>\n".format(section))
        # breadcrumb
        f.write("\t\t<nav aria-label='breadcrumb'>\n")
        f.write("\t\t<ol class='breadcrumb'>\n")
        f.write("\t\t<li class='breadcrumb-item'><a href='#'>Home</a></li>\n")
        f.write("\t\t<li class='breadcrumb-item'>Topics filed under{}</li>\n".format(section))
        f.write("\t\t</ol>\n")
        f.write("\t\t</nav>\n")
        # end breadcrumb
        
        i = 1
        topics = TopicCollector(section=section)
        for page in topics.scrap_topics_for_range_of_pages(start=start_page, stop=stop_page):

            f.write("\t\t\t<div id='js-scroll-target{}'>\n".format(i)) # div for targeting scroll
            f.write("\t\t\t<h2><a href='#js-scroll-target{0}' class='smooth-scroll'>Page {1}</a></h2>".format(i+1, i))
            i += 1
            
            for topic in list(page):
                f.write("\t\t\t<h3>Topic: <a href='{}' target='_blank'>{}</a></h3>".format(topic.url, topic.title))
                f.write("\t\t\t<h4>Poster: {}</h4>".format(topic.poster))
                f.write("\t\t\t<h5>{} <i class='fas fa-comment'></i></h5>".format(topic.comments))
                f.write("\t\t\t<div class='dropdown-divider' style='border:1px solid white;'></div>\n")
            f.write("\t\t\t</div>\n") # finish scroll div

        f.write("\t\t\t<p class='float-right'><a href='#topNav' class='smooth-scroll'>Back to top</a></p>\n")
        f.write("\t\t</div>\n")
        # jquery smooth scroll
        f.write("<script>\n")
        f.write("$(document).ready(function(){\n")
        f.write("\t $('.smooth-scroll').on('click', function(event) {\n")
        f.write("\t\tif (this.hash !== '') {\n")
        f.write("\t\t\tevent.preventDefault();\n")
        f.write("\t\t\tvar hash = this.hash;\n")
        f.write("\t\t\t$('html, body').animate({\n")
        f.write("\t\t\t scrollTop: $(hash).offset().top\n")
        f.write("\t\t\t}, 800, function(){\n")
        f.write("\t\t\t window.location.hash = hash;\n")
        f.write("\t\t\t});\n")
        f.write("\t\t}\n")
        f.write("\t});\n")
        f.write("});\n")
        f.write("</script>\n")
        # end jquery smooth scroll

        # finish up page structure
        f.write("\t</body>\n")
        f.write("</html>")
    print("Done hacking")
    os.startfile(destination_file)

if __name__ == "__main__":
    print("\n\nSee the associated Jupyter Notebook for usage instructions.")
