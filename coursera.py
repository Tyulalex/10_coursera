import requests
import logging
import time
import yaml
from xml.etree import ElementTree
from bs4 import BeautifulSoup
from openpyxl import Workbook


class Course:

    def __init__(self, raw_course_html, course_url):
        self.soup = BeautifulSoup(raw_course_html, 'html.parser')
        self.url = course_url

    @property
    def course_name(self):
        attribute = {'class': 'title display-3-text'}
        course_name_element = self.soup.find('h1', attrs=attribute)
        if course_name_element:
            return course_name_element.text

    @property
    def lang(self):
        attribute = {'class': 'language-info'}
        lang_element = self.soup.find('div', attrs=attribute)
        if lang_element:
            return lang_element.text

    @property
    def duration(self):
        commitment_element = self.soup.find('span', text='Commitment')
        if commitment_element:
            return commitment_element.parent.nextSibling.text

    @property
    def rating(self):
        rating_element_attr = {'class': 'ratings-text bt3-visible-xs'}
        rating_element = self.soup.find('div', attrs=rating_element_attr)
        if rating_element:
            return rating_element.text

    @property
    def start_date(self):
        start_date_attr = {
            'class': 'startdate rc-StartDateString caption-text'
        }
        start_date_element = self.soup.find('div', attrs=start_date_attr)
        if start_date_element:
            return start_date_element.text.replace(
                'Starts ', '').replace('Started ', '')


def fetch_data(url, attempts=2):
    sleep_time = 30
    response = requests.get(url, verify=False)
    if not response.ok:
        logger.warning(
            'response status code not ok {}'.format(response.status_code)
        )
        if attempts:
            time.sleep(sleep_time)
            return fetch_data(url, attempts=attempts - 1)
    return response


def load_config():
    with open('config.yaml') as config:
        return yaml.load(config)


def fetch_courses(course_data, namespace_mapping, courses_amount):
    root = ElementTree.fromstring(course_data)
    courses = list(
        map(lambda x: x.getchildren()[0].text,
            root.findall('urlset:url', namespace_mapping))
    )
    return courses[:courses_amount]


def write_course_row(ws, course):
    ws.append([
        course.course_name,
        course.start_date or 'N/A',
        course.lang or 'N/A',
        course.duration or 'N/A',
        course.rating or 'N/A',
        course.url
    ])


def write_course_column(ws):
    columns_name = (
        'Name',
        'Start Date',
        'Languages',
        'Duration',
        'Rating',
        'Url',
    )
    ws.append(columns_name)


def save_workbook(wb, filepath):
    wb.save(filepath)


def create_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Courses'
    return wb


def load_courses_data(courses_url):
    for course_url in courses_url:
        course_response = fetch_data(course_url)
        yield Course(course_response.text, course_url)


def fill_workbook(wb, courses_url):
    write_course_column(wb.active)
    courses_data = load_courses_data(courses_url)
    for course_data in courses_data:
        write_course_row(wb.active, course_data)


def get_logger():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    return logger

if __name__ == '__main__':
    logger = get_logger()
    config = load_config()
    courses_list = fetch_courses(
        course_data=fetch_data(config['courses_url']).text,
        namespace_mapping=config['namespace_mapping'],
        courses_amount=config['courses_amount']
    )
    if courses_list:
        wb = create_workbook()
        fill_workbook(wb, courses_list)
        save_workbook(wb, config['filepath'])
    logger.info('Script has finished it\'s work')
