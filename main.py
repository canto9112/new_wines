from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
import collections
from pprint import pprint


def get_first_wines(file):
    excel_data_df = pandas.read_excel(file, sheet_name='Лист1')
    excel_wine = excel_data_df.to_dict(orient='record')
    return excel_wine


def get_wines(file):
    excel_data_df = pandas.read_excel(file, sheet_name='Лист1')
    excel_wines = excel_data_df.to_dict(orient='record')
    wines = collections.defaultdict(list)
    for wine in excel_wines:
        category = wine['Категория']
        wines[category].append(wine)
    return wines


def work_years(year):
    years_now = datetime.datetime.now()
    years_work = years_now.year - year
    return years_work


def get_index_html(year, wines):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')

    rendered_page = template.render(
        years=year,
        wines=wines
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)


if __name__ == '__main__':
    foundation_year = 1920
    years_work = work_years(foundation_year)
    wines = get_wines('wine3.xlsx')

    get_index_html(years_work, wines)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()