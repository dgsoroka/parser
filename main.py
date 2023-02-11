import sys

import re
import xlsxwriter  # pip install XlsxWriter
import requests  # pip install requests
from bs4 import BeautifulSoup as bs  # pip install beautifulsoup4

headers = {'accept': '/',
           'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}
vacancy = input('Укажите название вакансии: ')
location = input('Укажите город поиска (Москва - 1, Владивосток - 22): ')
base_url = f'https://hh.ru/search/vacancy?area={location}&search_period=30&text={vacancy}&page='  # area=1 - Москва, search_period=3 - За 30 последних дня
pages = int(input('Укажите кол-во страниц для парсинга: '))
# Юрист+юрисконсульт

jobs = []


def hh_parse(base_url, headers):
    global end_with, start_with
    zero = 0
    while pages > zero:
        zero = str(zero)
        session = requests.Session()
        request = session.get(base_url + zero, headers=headers)
        if request.status_code == 200:
            soup = bs(request.content, 'html.parser')
            divs = soup.find_all('div', attrs={'data-qa': 'vacancy-serp__vacancy vacancy-serp__vacancy_standard'})
            for div in divs:
                title = div.find('a', attrs={'data-qa': 'serp-item__title'}).text
                compensation = ""
                # compensation = div.find('div', attrs={'data-qa': 'vacancy-salary-compensation-type-net'})
                # if compensation == None: # Если зарплата не указана
                #         compensation = 'None'
                # else:
                try:
                    compensation = div.find('span', attrs={'data-qa': 'vacancy-serp__vacancy-compensation'}).text
                    start_end = re.findall(r'\b\d+\d', compensation)
                    start_with = start_end[0] + start_end[1]
                except:
                    compensation = 'None'
                try:
                    end_with = start_end[2] + start_end[3]
                    end_with = int(end_with)
                except:
                    end_with = int(start_with)
                try:
                    href = div.find('a', attrs={'data-qa': 'serp-item__title'})['href']
                except:
                    href = "Error"
                try:
                    company = div.find('a', attrs={'data-qa': 'vacancy-serp__vacancy-employer'}).text
                except:
                    company = 'None'
                try:
                    text1 = div.find('div', attrs={'data-qa': 'vacancy-serp__vacancy_snippet_responsibility'}).text
                except:
                    text1 = 'None'
                try:
                    text2 = div.find('div', attrs={'data-qa': 'vacancy-serp__vacancy_snippet_requirement'}).text
                except:
                    text2 = 'None'
                content = text1 + '  ' + text2
                all_txt = [title, int(start_with), end_with, company, str(content), href]
                jobs.append(all_txt)
            # print(jobs)
            zero = int(zero)
            zero += 1

        else:
            print('error')
            zero = int(zero)

        # Запись в Excel файл
        workbook = xlsxwriter.Workbook('Vacancy.xlsx')
        worksheet = workbook.add_worksheet()
        # Добавим стили форматирования
        bold = workbook.add_format({'bold': 1})
        bold.set_align('center')
        center_H_V = workbook.add_format()
        center_H_V.set_align('center')
        center_H_V.set_align('vcenter')
        center_V = workbook.add_format()
        center_V.set_align('vcenter')
        cell_wrap = workbook.add_format()
        cell_wrap.set_text_wrap()

        # Настройка ширины колонок
        worksheet.set_column(0, 0, 35)  # A  https://xlsxwriter.readthedocs.io/worksheet.html#set_column
        worksheet.set_column(1, 1, 20)  # B
        worksheet.set_column(2, 2, 40)  # C
        worksheet.set_column(3, 3, 40)  # D
        worksheet.set_column(4, 4, 135)  # E

        worksheet.write('A1', 'Наименование', bold)
        worksheet.write('B1', 'Зарплата от', bold)
        worksheet.write('C1', 'Зарплата до', bold)
        worksheet.write('D1', 'Компания', bold)
        worksheet.write('E1', 'Описание', bold)
        worksheet.write('F1', 'Ссылка', bold)

        row = 1
        col = 0
        for i in jobs:
            worksheet.write_string(row, col, i[0], center_V)
            worksheet.write_number(row, col + 1, i[1], center_H_V)
            worksheet.write_number(row, col + 2, i[2], center_H_V)
            worksheet.write_string(row, col + 3, i[3], cell_wrap)
            # worksheet.write_url (row, col + 4, i[4], center_H_V)
            worksheet.write_url(row, col + 4, i[4], cell_wrap)
            worksheet.write_url(row, col + 5, i[5])
            row += 1

        print('OK')
    # print(jobs)
    workbook.close()


hh_parse(base_url, headers)
