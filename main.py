import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_products(file, table_sheet_name):
    excel_data_df = pandas.read_excel(file, sheet_name=table_sheet_name)
    excel_products = excel_data_df.to_dict(orient='record')
    products = collections.defaultdict(list)
    for product in excel_products:
        category = product['Категория']
        products[category].append(product)
    return products


def get_experience_of_work_years(year):
    current_year = datetime.datetime.now()
    return current_year.year - year


def fetch_index_html(experience_years, all_products, template_name):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template(template_name)
    rendered_page = template.render(
        years=experience_years,
        all_products=all_products
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)


if __name__ == '__main__':
    foundation_year = 1920
    products_file = 'wine3.xlsx'
    sheet_name = 'Лист1'
    template_name = 'template.html'

    experience_years = get_experience_of_work_years(foundation_year)
    products = get_products(products_file, sheet_name)

    fetch_index_html(experience_years, products, template_name)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()