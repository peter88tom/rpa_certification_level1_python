from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files

import time
import os

browser = Selenium()
url = "https://robotsparebinindustries.com"
user_name = "maria"
password = "thoushallnotpass"


def open_intranet_website():
    """ Opne the browser """
    browser.open_available_browser(url)


def log_in():
    """ Login maria using define credentials """
    usr = browser.find_element("//input[@name='username']")
    browser.input_text(usr, user_name)

    pwd = browser.find_element("//input[@name='password']")
    browser.input_text(pwd, password)

    login_button = browser.find_element("//button[@type='submit']")
    browser.click_button(login_button)


def download_the_excel_file():
    """ Download Excel file from remote server """
    http = HTTP()
    http.download(
        url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)


def fill_and_submit_the_form_for_one_person(sales_person):
    """ Fill in the sales form """
    # browser.wait_until_page_contains("css:sales-form", timeout=10)

    print(sales_person["First Name"])

    first_name = browser.find_element("//input[@name='firstname']")
    browser.input_text(first_name, sales_person["First Name"])

    last_name = browser.find_element("//input[@name='lastname']")
    browser.input_text(last_name, sales_person["Last Name"])

    sales_target = browser.find_element("//select[@id='salestarget']")
    browser.select_from_list_by_value(
        sales_target, f"{sales_person['Sales Target']}")

    sales_result = browser.find_element("//input[@name='salesresult']")
    browser.input_text(sales_result, sales_person["Sales"])

    submit_sale = browser.find_element("//button[@type='submit']")
    browser.click_button(submit_sale)

    # time.sleep(30)


def fill_the_form_using_the_data_from_the_excel_file():
    time.sleep(5)
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    sales_persons = excel.read_worksheet_as_table(header=True)
    excel.close_workbook()

    for sales_person in sales_persons:
        fill_and_submit_the_form_for_one_person(sales_person)


def collect_the_results():
    pass


def export_the_table_as_pdf():
    pass


def log_out():
    pass


def main():
    try:
        open_intranet_website()
        log_in()
        download_the_excel_file()
        fill_the_form_using_the_data_from_the_excel_file()
        collect_the_results()
        export_the_table_as_pdf()
    finally:
        browser.close_browser()


if __name__ == "__main__":
    main()
