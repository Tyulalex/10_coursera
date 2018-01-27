import requests
import logging
import time
import yaml
from xml.etree import ElementTree
from bs4 import BeautifulSoup
from openpyxl import Workbook


class Course:
    def __init__(self, req_response):
        self.soup = BeautifulSoup(req_response.text, 'html.parser')
        self.url = req_response.url

    @property
    def course_name(self):
        attribute = {"class": "title display-3-text"}
        return getattr(
            self.soup.find("h1", attrs=attribute), 'text', '-').encode('utf-8')

    @property
    def lang(self):
        attribute = {"class": "language-info"}
        return getattr(self.soup.find("div", attrs=attribute), 'text', '-')

    @property
    def duration(self):
        commitment_element = self.soup.find("span", text="Commitment")
        if commitment_element:
            return commitment_element.parent.nextSibling.text
        return '-'

    @property
    def rating(self):
        rating_element_attr = {"class": "ratings-text bt3-visible-xs"}
        return getattr(
            self.soup.find("div", attrs=rating_element_attr), 'text', '-'
        )

    @property
    def start_date(self):
        start_date_attr = {
            "class": "startdate rc-StartDateString caption-text"
        }
        return getattr(
            self.soup.find("div", attrs=start_date_attr),
            'text',
            '-'
        ).replace("Starts ", "").\
            replace("Started ", "").\
            encode('utf-8')


def send_get_request(url, attempts=2):
    sleep_time = 30
    response = requests.get(url, verify=False)
    if not response.ok:
        logger.warning(
            "response status code not ok {}".format(response.status_code)
        )
        if attempts:
            time.sleep(sleep_time)
            return send_get_request(url, attempts=attempts - 1)
    return response


def load_config():
    with open('config.yaml') as config:
        return yaml.load(config)


def filter_courses(course_data, namespace_mapping, courses_amount):
    root = ElementTree.fromstring(course_data)
    courses = list(
        map(lambda x: x.getchildren()[0].text,
            root.findall('urlset:url', namespace_mapping))
    )
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
        'Name',
        'Start Date',
        'Languages',
        'Duration',
        'Rating',
        'Url'
    )
    for i, col in enumerate(columns_name):
        ws.cell(column=i + 1, row=1, value=col)


def save_workbook(wb, filepath):
    wb.save(filepath)


def create_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Courses'
    return wb


def load_courses_data(courses_url):
    for i, course_url in enumerate(courses_url):
        yield Course(send_get_request(course_url))


def fill_workbook(wb, courses_url):
    write_course_column(wb.active)
    courses_data = load_courses_data(courses_url)
    for i, course_data in enumerate(courses_data):
        write_course_row(wb.active, i + 2, course_data)


def get_logger():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    return logger

if __name__ == '__main__':
    logger = get_logger()
    config = load_config()
    courses_list = filter_courses(
        course_data=send_get_request(config['courses_url']).text,
        namespace_mapping=config['namespace_mapping'],
        courses_amount=config['courses_amount']
    )
    if courses_list:
        wb = create_workbook()
        fill_workbook(wb, courses_list)
        save_workbook(wb, config['filepath'])
    logger.info("Script has finished it's work")
