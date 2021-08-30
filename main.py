from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import time
from openpyxl import Workbook
import os
import requests
from selenium.webdriver.chrome.options import Options
import sys

current_dir = os.path.abspath(os.getcwd())

if not os.path.exists('output'):
    os.makedirs('output')


wb = Workbook()
dest_file_name = 'output/result.xlsx'
ws1 = wb.active
ws1.title = 'Agencies'
ws1['A' + '1'] = 'Agency name'
ws1['B' + '1'] = 'Spend Amounts'
ws2 = wb.create_sheet(title=' individual investments')

options = Options()
options.add_experimental_option('prefs', {
    "download.default_directory": os.path.join(current_dir, 'output/'),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
}
)
driver = webdriver.Chrome(options=options)
url = 'https://itdashboard.gov/'
driver.get(url)
agencies = {}
links_to_visit = []

def is_element_present_on_page(element_xpath, count=5):
    try:
        return driver.find_element_by_xpath(element_xpath)
    except NoSuchElementException:
        time.sleep(3)
        while count != 0:
            count -= 1
            return is_element_present_on_page(element_xpath, count)
        else:
            return None

def run():
    dive_in_button = driver.find_element_by_xpath('/html/body/main/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/a')
    dive_in_button.click()
    if is_element_present_on_page('//*[@id="agency-tiles-widget"]/div/div[1]/div[1]/div/div/div/div[1]/a/span[1]'):
        agencies_elements = driver.find_elements_by_xpath('//*[@id="agency-tiles-widget"]/div/div/div/div/div/div/div[1]/a')
        for item in agencies_elements:
            name = item.find_element_by_xpath('./span[1]').text
            value = item.find_element_by_xpath('./span[2]').text
            agencies[name] = value


def write_xlsx():
    cell_no = 2
    for key, value in agencies.items():
        ws1['A' + str(cell_no)] = key
        ws1['B' + str(cell_no)] = value
        cell_no += 1


def check_agency(agency_name):
    agencies_elements = driver.find_elements_by_xpath('//*[@id="agency-tiles-widget"]/div/div/div/div/div/div/div[1]/a')
    for item in agencies_elements:
        name = item.find_element_by_xpath('./span[1]').text
        if name == agency_name:
            item.click()
            break

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    if is_element_present_on_page('//*[@id="investments-table-object_wrapper"]/div[3]/div[1]/div/table/thead/tr[2]/th[1]', 10):
        driver.find_element_by_xpath('//*[@id="investments-table-object_length"]/label/select').click()
        driver.find_element_by_xpath('//*[@id="investments-table-object_length"]/label/select/option[4]').click()
        parse_table()

def is_loading(tries=10):
    try:
        driver.find_element_by_class_name('loading')
        tries -= 1
        time.sleep(3)
        if tries > 0:
            return is_loading(tries)
        else:
            return True
    except:
        return False


def parse_table():
    if not is_loading():
        table_body = driver.find_element_by_xpath('//*[@id="investments-table-object"]/tbody')
        for row in table_body.find_elements_by_xpath('./tr'):
            _row = []
            for index, col in enumerate(row.find_elements_by_xpath('./td')):
                _row.append(col.text)
                link_text = col.find_elements_by_link_text(col.text)
                if link_text:
                    links_to_visit.append(link_text[0].get_attribute('href'))
            ws2.append(_row)


# def download_pdf_files_from_links():
#       Download via requests
#     for item in links_to_visit:
#         r = requests.get(item)
#         file_name = f'{str(item).split("/").pop()}.pdf'
#         with open(os.path.join(current_dir, 'output/', file_name), 'wb') as f:
#             f.write(r.content)

def download_pdf_files_from_links():
    # Download via selenium
    for item in links_to_visit:
        driver.get(item)
        if is_element_present_on_page('//*[@id="business-case-pdf"]/a'):
            driver.find_element_by_xpath('//*[@id="business-case-pdf"]/a').click()
            time.sleep(10)


if __name__ == '__main__':
    agency_name = sys.argv.pop()
    run()
    write_xlsx()
    check_agency(agency_name)
    wb.save(filename=dest_file_name)
    download_pdf_files_from_links()
    driver.close()
