from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
from pprint import pprint


def get_wines(file):
    excel_data_df = pandas.read_excel(file, sheet_name='list1')
    excel_wine = excel_data_df.to_dict(orient='record')
    return excel_wine


def work_years(year):
    years_now = datetime.datetime.now()
    years_work = years_now.year - year
    return years_work


def get_years_index_html(year):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')
    wines = get_wines('wine.xlsx')
    rendered_page = template.render(
        years=year,
        wines=wines
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)


# wines = get_wines('wine.xlsx')
# pprint(wines)
# for wine in wines:
#     wine_name = wine['Название']
#     sort_wine = wine['Сорт']
#     price_wine = wine['Цена']
#     image_wine = wine['Картинка']


if __name__ == '__main__':
    foundation_year = 1920
    years_work = work_years(foundation_year)

    get_years_index_html(years_work)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
