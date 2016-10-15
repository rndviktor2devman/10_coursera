import requests
import random
import json
from xml.etree import ElementTree as ETree
from bs4 import BeautifulSoup as BS
from openpyxl import Workbook


def get_courses_list(output_length):
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    prefix = '{http://www.sitemaps.org/schemas/sitemap/0.9}'
    urls_text = requests.get(url).content
    tree_root = ETree.fromstring(urls_text.decode('utf-8'))
    xml_courses = tree_root.findall('{}url'.format(prefix))

    links_list = []
    for course in random.sample(xml_courses, output_length):
        links_list.append(course.find('{}loc'.format(prefix)).text)
    return links_list


def get_course_info(course_slug):
    course_data = {}
    try:
        r = requests.get(course_slug)
    except requests.exceptions.ConnectionError:
        return get_course_info(course_slug)
    soup = BS(r.content, "lxml")
    try:
        course_data['title'] = soup.find('div', "title display-3-text").string
    except AttributeError:
        course_data['title'] = ''
    try:
        for span in soup.find_all('div', "ratings-text bt3-hidden-xs")[0]:
            if isinstance(span, str) and span.startswith("Average"):
                course_data['rating'] = span.replace("Average User Rating ", "")
    except (IndexError, AttributeError):
        course_data['rating'] = ''
    try:
        course_data['weeks'] = soup.find_all(
            'div', "week-heading body-2-text"
        )[-1].string
    except IndexError:
        course_data['weeks'] = ''
    try:
        course_data['date'] = json.loads(
            soup.find('div', "rc-CourseGoogleSchemaMarkup").script.text
        )['hasCourseInstance'][0]['startDate']
        course_data['language'] = json.loads(
            soup.find('div', "rc-CourseGoogleSchemaMarkup").script.text
        )['hasCourseInstance'][0]['inLanguage']
    except (AttributeError, KeyError):
        course_data['date'] = ''
        course_data['language'] = ''

    course_data['url'] = course_slug
    return course_data


def output_courses_info_to_xlsx(filepath, data_set):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Coursera Courses"
    worksheet.append(
        ["Title",
         "URL",
         "Language",
         "StartDate",
         "WeeksCount",
         "AverageRating"])

    for info in data_set:
        if info is not None:
            worksheet.append(
                [info['title'],
                 info['url'],
                 info['language'],
                 info['date'],
                 info['weeks'],
                 info['rating']]
            )
    workbook.save(filepath)


if __name__ == '__main__':
    list_links = get_courses_list(20)
    file_path = "output.xlsx"
    data_set = []
    for link in list_links:
        data_set.append(get_course_info(link))
    output_courses_info_to_xlsx(file_path, data_set)
