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


def get_second_wines(file):
    excel_data_df = pandas.read_excel(file, sheet_name='Лист1')
    excel_wines = excel_data_df.to_dict(orient='record')
    wines = collections.defaultdict(list)
    for wine in excel_wines:
        category = wine['Категория']
        wines[category].append(wine)
    categories = []
    products = []
    for key, value in wines.items():
        categories.append(key)
        products.append(value)
    products_numbers = len(products[:])
    return categories, products[:], products_numbers


def work_years(year):
    years_now = datetime.datetime.now()
    years_work = years_now.year - year
    return years_work


def get_index_html(year, categories, products, products_numbers):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        years=year,
        categories=categories,
        products=products[products_numbers]
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)


if __name__ == '__main__':
    foundation_year = 1920
    years_work = work_years(foundation_year)
    categories, products, products_numbers = get_second_wines('wine2.xlsx')

    for number in range(0, products_numbers):
        print(number)
        get_index_html(years_work, categories, products, number)
        break

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()