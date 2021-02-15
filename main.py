from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
from pprint import pprint


excel_data_df = pandas.read_excel('wine.xlsx', sheet_name='list1')
excel_wine = excel_data_df.to_dict(orient='record')

pprint(excel_wine)


def work_years(year):
    years_now = datetime.datetime.now()
    years_work = years_now.year - year
    return years_work


def get_index_html(year):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    rendered_page = template.render(
        years=year,
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)


if __name__ == '__main__':
    foundation_year = 1920
    years_work = work_years(foundation_year)

    get_index_html(years_work)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
