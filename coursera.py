import requests
import logging
import time
import yaml
from xml.etree import ElementTree
from bs4 import BeautifulSoup
from openpyxl import Workbook


class Course:
    def __init__(self, url):
        self.soup = BeautifulSoup(fetch_data(url), 'html.parser')
        self.url = url

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


def fetch_data(url, default_response=None):
    attempt = 0
    response = requests.get(url, verify=False)
    if not response.ok:
        logger.warning(
            "response status code not ok {}".format(
                response.status_code))
        attempt += 1
        if attempt:
            time.sleep(30)
            return fetch_data(url)
        return default_response
    return response.text


def load_config():
    with open('config.yaml') as config:
        return yaml.load(config)


def load_courses(courses_url):
    return fetch_data(courses_url, default_response=[])


def filter_courses(course_data, namespace_mapping, courses_amount):
    root = ElementTree.fromstring(course_data)
    courses = list(map(lambda x: x.getchildren()[0].text,
                       root.findall('urlset:url', namespace_mapping)))
    return courses[:courses_amount]


def write_course_row(ws, row, course):
    ws.cell(row=row, column=1, value=course.course_name)
    ws.cell(row=row, column=2, value=course.start_date)
    ws.cell(row=row, column=3, value=course.lang)
    ws.cell(row=row, column=4, value=course.duration)
    ws.cell(row=row, column=5, value=course.rating)
    ws.cell(row=row, column=6, value=course.url)


def write_course_column(ws):
    columns_name = (
        'Name', 'Start Date', 'Languages', 'Duration', 'Rating', 'Url')
    for col in range(1, len(columns_name) + 1):
        ws.cell(column=col, row=1, value=columns_name[col - 1])


def save_workbook(wb, filepath):
    wb.save(filepath)


def create_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Courses'
    return wb


def fill_workbook(wb, courses_url):
    write_course_column(wb.active)
    for i, course_url in enumerate(courses_url):
        course = Course(course_url)
        write_course_row(wb.active, i + 2, course)


def get_logger():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    return logger

if __name__ == '__main__':
    logger = get_logger()
    config = load_config()
    courses_list = filter_courses(load_courses(config['courses_url']),
                                  config['namespace_mapping'],
                                  config['courses_amount'])
    if courses_list:
        wb = create_workbook()
        fill_workbook(wb, courses_list)
        save_workbook(wb, config['filepath'])
    logger.info("Script has finished it's work")
