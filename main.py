import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_products(file, sheet_name):
    excel_data_df = pandas.read_excel(file, sheet_name=sheet_name)
    excel_products = excel_data_df.to_dict(orient='record')
    products = collections.defaultdict(list)
    for product in excel_products:
        category = product['Категория']
        products[category].append(product)
    return products


def get_work_years(year):
    years_now = datetime.datetime.now()
    work_years = years_now.year - year
    return work_years


def rendered_index_html(year, wines):
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
    products_file = 'wine3.xlsx'
    sheet_name = 'Лист1'

    years_work = get_work_years(foundation_year)
    wines = get_products(products_file, sheet_name)

    rendered_index_html(years_work, wines)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()