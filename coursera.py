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
    r = requests.get(course_slug)
    soup = BS(r.content, "lxml")
    empty_item = 'not found'

    title_tag = soup.find('div', "title display-3-text")
    course_data['title'] = empty_item
    if title_tag is not None:
        course_data['title'] = title_tag.text

    rating_tag = soup.find('div', {'class': 'ratings-text bt3-visible-xs'})
    course_data['rating'] = 'not found'
    if rating_tag is not None:
        course_data['rating'] = rating_tag.text

    week_tag = soup.find('div', {'class': 'rc-WeekView'})
    course_data['weeks'] = empty_item
    if week_tag is not None:
        course_data['weeks'] = len(
            week_tag.find_all('div', {'class': 'week'}))

    info_tag = soup.find('script', {'type': 'application/ld+json'})
    course_data['date'] = empty_item
    if info_tag is not None:
        json_info = json.loads(info_tag.text)
        if 'startDate' in json_info['hasCourseInstance'][0].keys():
            course_data['date'] = json_info['hasCourseInstance'][0]['startDate']

    language_tag = soup.find('div', {'class': 'language-info'})
    course_data['language'] = empty_item
    if language_tag is not None:
        course_data['language'] = language_tag.text

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
