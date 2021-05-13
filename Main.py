from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import requests
import pandas as pd
import xlwings as xw

HEADERS = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36',
    'accept': '*/*'}

HOST = 'https://www.avangard.ru/rus/about/office/moskva/'


# Get source from site with js elements
def get_page_source():
    driver_exe = 'chromedriver'
    options = Options()
    options.add_argument("--headless")  # hide chromedriver
    driver = webdriver.Chrome(driver_exe, options=options)
    driver.get(HOST)
    driver.find_element_by_link_text('В другом городе').click()
    time.sleep(2)
    source = driver.page_source
    return source


# Create file
def create_xlsx():
    try:
        wb = xw.Book(r'C:\Users\korni\Documents\Python Projects\Avang\Offices.xlsx')  # Set direction !!!
        sheet = wb.sheets[0]
    except FileNotFoundError:
        wb = xw.Book()
        wb.save(r'C:\Users\korni\Documents\Python Projects\Avang\Offices.xlsx')
        wb = xw.Book(r'C:\Users\korni\Documents\Python Projects\Avang\Offices.xlsx')  # Set direction !!!
        sheet = wb.sheets['Sheet1']
    return sheet


# Get list include name of city and link
def get_city_list(source):
    soup = BeautifulSoup(source, 'html.parser')
    city_name = soup.find('div', class_='cityHolder').find('br').find_all_next('li')
    cities = []
    for city in city_name:
        cities.append({
            'name': city.find('a').get_text(strip=True),
            'link': HOST + city.find('a').get('href')
        })
    return cities


# Get html code from link
def get_html(url, params=None):
    try:
        r = requests.get(url, headers=HEADERS, params=params, timeout=10)
        return r
    except requests.exceptions.Timeout:
        print('Timeout occurred')
        return 0


# Get content from html
def get_content(html):
    soup = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('tr', class_='shows')
    office = []
    for item in items:
        office.append({
            '№ Офиса': item.find('td', class_='name').get_text(strip=True),
            'Адрес': item.find('td', class_='address').find_next('span').get_text(strip=True),
            'Метро': item.find('td', class_='address').find_next('div').get_text(' ', strip=True),
            'Телефон': item.find('td', class_='phone').get_text(strip=True),
            'Время работы': item.find('td', class_='timeWork').get_text(strip=True)
        })
    return office


# Main function
def parse():
    source = get_page_source()
    cities = get_city_list(source)  # Cities dictionary
    sheet = create_xlsx()
    iteration = 0
    while True:
        df = pd.DataFrame()
        offices = []
        for url in cities:  # Get link from the list of cities
            print('Анализ города' ' ' + (url.get('name')))
            offices.append({'Ссылка на карту города': url.get('link')})
            url = url.get('link')
            html = get_html(url)
            if html == 0:
                df = pd.DataFrame()
                break
            if html.status_code == 200:
                offices.extend(get_content(html.text))
                df = pd.DataFrame(offices)
        iteration += 1
        if df.empty:
            print('Выход')
            print('Проход №', iteration, 'завершен неуспешно')
            continue
        else:
            sheet.range('A1').value = df
            print('Проход №', iteration, 'завершен успешно')
        # time.sleep(180)


parse()
