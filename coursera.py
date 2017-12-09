import requests
import logging
import yaml
from xml.etree import ElementTree
from bs4 import BeautifulSoup
from openpyxl import Workbook


class Course:
    def __init__(self, url):
        self.url = url
        self.soup = self._get_content()

    def _get_content(self):
        course_html = requests.get(self.url, verify=False)
        return BeautifulSoup(course_html.text, 'html.parser')

    @property
    def course_name(self):
        attribute = {"class": "title display-3-text"}
        return getattr(self.soup.find("h1", attrs=attribute),
                       'text', '-').encode('utf-8')

    @property
    def lang(self):
        attribute = {"class": "language-info"}
        return getattr(self.soup.find("div", attrs=attribute),
                       'text', '-')

    @property
    def duration(self):
        commitment_element = self.soup.find(
            "span", text="Commitment")
        if commitment_element:
            return commitment_element.parent.\
                nextSibling.text
        return '-'

    @property
    def rating(self):
        rating_element_attr = {
            "class": "ratings-text bt3-visible-xs"
        }
        return getattr(
            self.soup.find(
                "div", attrs=rating_element_attr), 'text', '-')

    @property
    def start_date(self):
        start_date_attr = {
            "class": "startdate rc-StartDateString caption-text"
        }
        return getattr(
            self.soup.find(
                "div", attrs=start_date_attr), 'text', '-').\
            replace("Starts ", "").\
            replace("Started ", "").encode('utf-8')


def load_config():
    with open('config.yaml') as config:
        return yaml.load(config)


def get_courses_list(courses_url, namespace_mapping):
    response = requests.get(courses_url, verify=False)
    if response.status_code != requests.codes.ok:
        logger.error(
            'ERROR status code not OK: {}'.format(
                response.status_code))
        return []
    root = ElementTree.fromstring(response.text)
    courses = list(map(lambda x: x.getchildren()[0].text,
                       root.findall('urlset:url', namespace_mapping)))
    return courses[:20] if len(courses) > 20 else courses


def write_course_row(ws, row, course):
    ws.cell(row=row, column=1, value=course.course_name)
    ws.cell(row=row, column=2, value=course.start_date)
    ws.cell(row=row, column=3, value=course.lang)
    ws.cell(row=row, column=4, value=course.duration)
    ws.cell(row=row, column=5, value=course.rating)
    ws.cell(row=row, column=6, value=course.url)


def output_courses_info_to_xlsx(filepath, courses_url):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Courses'
    columns_name = (
        'Name', 'Start Date', 'Languages', 'Duration', 'Rating', 'Url')
    for col in range(1, len(columns_name) + 1):
        ws.cell(column=col, row=1, value=columns_name[col-1])
    for i, course_url in enumerate(courses_url):
        course = Course(course_url)
        write_course_row(ws, i + 2, course)
    wb.save(filepath)


def get_logger():
    logger = logging.getLogger('courseraapp')
    hdlr = logging.FileHandler('courseraapp.log', encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
    hdlr.setFormatter(formatter)
    logger.addHandler(hdlr)
    logger.setLevel(logging.DEBUG)
    return logger

if __name__ == '__main__':
    logger = get_logger()
    config = load_config()
    courses_list = get_courses_list(config['courses_url'],
                                    config['namespace_mapping'])
    if courses_list:
        output_courses_info_to_xlsx(config['filepath'], courses_list)
    logger.info("Script has finished it's work")
