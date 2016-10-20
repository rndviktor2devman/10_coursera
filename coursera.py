import requests
import random
import json
from xml.etree import ElementTree as ETree
from bs4 import BeautifulSoup as BS
from openpyxl import Workbook


def get_courses_list(output_length):
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    tree_root = ETree.fromstring(requests.get(url).content)
    list_urls = [child[0].text for child in tree_root]
    random.shuffle(list_urls)
    return list_urls[:output_length]


def get_course_info(course_slug):
    course_data = {}
    try:
        r = requests.get(course_slug)
    except requests.exceptions.ConnectionError:
        return None
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
    for link in list_links:
        print(link)
    file_path = "output.xlsx"
    data_set = []
    for link in list_links:
        data_set.append(get_course_info(link))
    output_courses_info_to_xlsx(file_path, data_set)
