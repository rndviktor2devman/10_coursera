import requests
import random
import json
from xml.etree import ElementTree as ETree
from bs4 import BeautifulSoup as BS
from openpyxl import Workbook


def get_courses_list():
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    tree_root = ETree.fromstring(requests.get(url).content)
    list_urls = [child[0].text for child in tree_root]
    random.shuffle(list_urls)
    return list_urls


def get_course_info(course_slug):
    course_data = {}
    try:
        r = requests.get(course_slug)
        soup = BS(r.content, "lxml")
        course_data['title'] = soup.find('div', "title display-3-text").string
        for span in soup.find_all('div', "ratings-text bt3-hidden-xs")[0]:
            if isinstance(span, str) and span.startswith("Average"):
                course_data['rating'] = span.replace("Average User Rating ", "")
        course_data['weeks'] = soup.find_all(
            'div', "week-heading body-2-text"
        )[-1].string
        course_data['date'] = json.loads(
            soup.find('div', "rc-CourseGoogleSchemaMarkup").script.text
        )['hasCourseInstance'][0]['startDate']
        course_data['language'] = json.loads(
            soup.find('div', "rc-CourseGoogleSchemaMarkup").script.text
        )['hasCourseInstance'][0]['inLanguage']
    except (requests.exceptions.ConnectionError, IndexError, AttributeError, KeyError):
        return None

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
    output_length = 20
    list_links = get_courses_list()
    file_path = "output.xlsx"
    data_set = []
    number_link = 0
    while number_link < len(list_links) and len(data_set) < output_length:
        course_data = get_course_info(list_links[number_link])
        if course_data is not None:
            data_set.append(course_data)
            print(course_data)

        number_link += 1
    output_courses_info_to_xlsx(file_path, data_set)
