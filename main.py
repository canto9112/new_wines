import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_products(filename, table_sheet_name):
    excel_data_df = pandas.read_excel(filename, sheet_name=table_sheet_name)
    excel_products = excel_data_df.to_dict(orient='record')
    products = collections.defaultdict(list)
    for product in excel_products:
        category = product['Категория']
        products[category].append(product)
    sorted_products = {key: value for key, value in sorted(products.items())}
    return sorted_products


def get_experience_of_work_years(year):
    current_year = datetime.datetime.now()
    experience_years = current_year.year - year
    last_digit = experience_years % 10
    two_last_digit = abs(experience_years) % 100
    if last_digit == 1:
        suffix = "год"
        if two_last_digit == 11:
            suffix = "лет"
    elif 2 <= last_digit <= 4:
        suffix = "года"
        if 11 <= two_last_digit <= 19:
            suffix = "лет"
    else:
        suffix = "лет"
    experience_years = f"{experience_years} {suffix}"
    return experience_years


def get_template(path, template_name):
    env = Environment(
        loader=FileSystemLoader(path),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template(template_name)
    return template


def fetch_rendered_page(template, experience_years, all_products):
    rendered_page = template.render(
        years=experience_years,
        all_products=all_products
    )
    return rendered_page


def fetch_index_html(homepage, rendered_page):
    with open(homepage, 'w', encoding="utf8") as file:
        file.write(rendered_page)


def main():
    foundation_year = 1920
    filename = 'wine3.xlsx'
    sheet_name = 'Лист1'
    template_name = 'template.html'
    homepage_name = 'index.html'
    filepath = '.'

    experience_years = get_experience_of_work_years(foundation_year)
    products = get_products(filename, sheet_name)

    template = get_template(filepath, template_name)

    rendered_page = fetch_rendered_page(template, experience_years, products)

    fetch_index_html(homepage_name, rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()