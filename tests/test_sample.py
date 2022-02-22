import os
from asyncio.log import logger
from time import sleep
from RPA.Tables import Tables
from RPA.Browser.Selenium import Selenium
from ExcelFileExtend import FilesExtend
from html_tables import read_table_from_html

# When importing from RPA, create a instance of that class
sel = Selenium()
excel = FilesExtend()
tables = Tables()

CURDIR                    = os.path.dirname(__file__)
data_file                 = f'{CURDIR}/ListItems.xlsx'
URL                       = 'http://127.0.0.1:8000/accounts/login/'
username_element          = '//input[@name="username"] '
password_element          = '//input[@name="password"]'
login_button              = '//button[@type="submit"]'
product_code_element      = '//input[@id="id_code"]'
product_name_element      = '//input[@id="id_name"]'
quantity_element          = '//input[@id="id_quantity"]'
unit_element              = '//input[@id="id_unit"]'
unit_price_element        = '//input[@id="id_unit_price"]'
total_amount_element      = '//input[@id="id_total_amount"]'
submit_button             = '//button[@type="submit"]'
create_new_product_button = '//*[text()="Create product"]'
logout_element            = '//a[@href="/accounts/logout/"]'
back_to_main_element      = '//a[@href="/"]'
# variable for query items'
table_element             = '//table[@class="mt-2 table"]'
Result_file               = f'{CURDIR}/Result.xlsx'
header_table = '<h2>List products need to refill\n</h2>\
<table style="border: 1px solid black;border-collapse: collapse;width:100%"><tr>\
<th style="border: 1px solid black;border-collapse: collapse">Product Code</th>\
<th style="border: 1px solid black;border-collapse: collapse">Product Name</th>\
<th style="border: 1px solid black;border-collapse: collapse">Quantity</th>\
<th style="border: 1px solid black;border-collapse: collapse">Unit</th>\
<th style="border: 1px solid black;border-collapse: collapse">Unit Price</th>\
<th style="border: 1px solid black;border-collapse: collapse">Total Amount</th></tr><tr>'


def read_data_to_table(file, sheet, header=True):
    excel.open_workbook_with_data(file)
    worksheet = excel.read_worksheet(name=sheet, header=header)
    order = tables.create_table(worksheet, trim=True)
    excel.close_workbook()
    return order


def submit_data_to_app(data):
    sel.input_text_when_element_is_visible(product_code_element, data['Product code'])
    sel.input_text(product_name_element, data['Product Name'])
    sel.input_text(quantity_element, data['Quantity'])
    sel.input_text(unit_element, data['Unit'])
    sel.input_text(unit_price_element, data['Unit Price'])
    sel.input_text(total_amount_element, data['Total Amount'])
    highlight(submit_button)
    sel.click_button(submit_button)


def show_result_and_logout_the_app():
    highlight(logout_element)
    sel.click_element(logout_element)


def login_the_app_and_choose_add_new_item():
    sel.maximize_browser_window()
    sel.input_text(username_element, 'admin')
    sel.input_text(password_element, 'admin')
    highlight(submit_button)
    sel.click_element(login_button)
    highlight(create_new_product_button)
    sel.click_element_when_visible(create_new_product_button)


def get_html_table():
    highlight(table_element)
    html_table = sel.get_element_attribute(table_element, 'outerHTML')
    return html_table


def read_html_table_as_table():
    html_table = get_html_table()
    table = read_table_from_html(html_table)
    return table


def highlight(location):
    sel.highlight_elements(location, style='solid', width='2px', color='blue')
    sleep(0.5)
    sel.clear_all_highlights()


def create_new_workbook_from_map(data):
    input_values = {
        'ProductCode': data[0],
        'ProductName': data[1],
        'Quantity': data[2],
        'Unit': data[3],
        'Unit Price': data[4],
        'TotalAmount': data[5],
    }
    excel.append_rows_to_worksheet(input_values, header=True)


def create_row_in_table_on_report(data):
    text_value = f'<td style="border: 1px solid black;border-collapse: collapse">{data[0]}</td>'
    text_value = f'{text_value}<td style="border: 1px solid black;border-collapse: collapse">{data[1]}</td>'
    text_value = f'{text_value}<td style="border: 1px solid black;border-collapse: collapse">{data[2]}</td>'
    text_value = f'{text_value}<td style="border: 1px solid black;border-collapse: collapse">{data[3]}</td>'
    text_value = f'{text_value}<td style="border: 1px solid black;border-collapse: collapse">{data[4]}</td>'
    text_value = f'{text_value}<td style="border: 1px solid black;border-collapse: collapse">{data[5]}</td>'
    return text_value


# def create_body_in_table_on_report(data):
# Not used


# def read_data_to_table(file, sheet, header=True):
# Defined above. This is dupplicated code


def send_report(body):
    result = ''
    for content in body:
        result = f'{result}{content}'


def query_items_and_send_report():
    highlight('//a[@href="/"]')
    sel.click_link('//a[@href="/"]')
    data_table = read_html_table_as_table()
    dimensions = tables.get_table_dimensions(data_table)
    excel.create_workbook(Result_file)
    content_html = [header_table]
    for index in range(dimensions[0]):
        row = tables.get_table_row(data_table, index, as_list=True)
        quantity = row[2]
        quantity = float(quantity)
        if quantity < 500.0:
            create_new_workbook_from_map(row)
            temp = create_row_in_table_on_report(row)
            content_html.append(f'<tr>{temp}</tr>')
    excel.save_workbook()
    # send_report(content_html)


def minimal_task():
    sel.open_chrome_browser(URL, headless=True)
    login_the_app_and_choose_add_new_item()
    item_list = read_data_to_table(data_file, 'Sheet1')
    for item in item_list:
        submit_data_to_app(item)
    query_items_and_send_report()
    show_result_and_logout_the_app()
    logger.info('Done.')


if __name__ == "__main__":
    minimal_task()
    sel.close_browser()
