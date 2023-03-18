import pandas as pd
import requests
import time
import os
import PyPDF4
from bs4 import BeautifulSoup
from datetime import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

def load_data(file_path):
    df = pd.read_excel(file_path)
    df['ИНН'] = df['ИНН'].fillna('')
    return df

def get_info_from_site(inn):
    url = f"https://bankruptcy.kommersant.ru/search/index.php?query={inn}"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    ads = soup.find_all('div', {'class': 'page-content-company'})
    return ads, url

def save_pdf_file(row, ad, pdf_folder_path, url):
    file_name = f"Siebel-DI_{row['Фамилия']}_{row['Имя']}_{row['Отчество']}_{ad+1}.pdf"
    pdf_file_path = os.path.join(pdf_folder_path, file_name)

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    try:
        driver.get(url)
        print_button = driver.find_element_by_class_name("js-print-article")
        print_button.click()
        time.sleep(5)
        save_as_pdf_button = driver.find_element_by_xpath("//span[@class='action-name' and text()='Сохранить в PDF']")
        save_as_pdf_button.click()
        time.sleep(5)
        driver.quit()
        return f"Файл {pdf_file_path} сохранен"
    except:
        return f"Ошибка при сохранении файла {pdf_file_path}"

def process_data(df, pdf_folder_path):
    df['Найдено в Ъ'] = ''
    for index, row in df.iterrows():
        inn = row['ИНН']
        if inn:
            ads, url = get_info_from_site(inn)
            if ads:
                df.at[index, 'Найдено в Ъ'] = 'Найдена информация об ИНН'
                for i, ad in enumerate(ads):
                    save_pdf_file(row, i, pdf_folder_path, url)

def main(file_path, pdf_folder_path):
    df = load_data(file_path)
    process_data(df, pdf_folder_path)

main('xlsx/сlients.xlsx', 'pdf')

