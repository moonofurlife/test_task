import pandas as pd
import requests
import os
import time
import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

# Функция для загрузки данных из excel файла
def load_data(file_path):
    df = pd.read_excel(file_path)
    # Если в колонке 7 есть пропущенные значения, заполним их пустыми строками
    df.iloc[:, 8] = df.iloc[:, 8].fillna('')
    # Добавляем две новые колонки: "Найдено в Ъ" и "Дата проверки"
    df["Найдено в Ъ"] = ""
    df["Дата проверки"] = ""
    return df

# Функция для поиска информации на сайте и сохранения ее в pdf файл
def search_and_save_pdf(inn, pdf_folder_path):
    url = f"https://bankruptcy.kommersant.ru/search/index.php?query={inn}"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    # Ищем все объявления, соответствующие запросу
    ads = soup.find_all('div', {'class': 'page-content-company'})
    # Если объявления найдены
    if ads:
        # date_str = datetime.datetime.now().strftime("%Y-%m-%d")
        # pdf_folder_path = os.path.join(pdf_folder_path, date_str)
        # os.makedirs(pdf_folder_path, exist_ok=True)
        # # driver = webdriver.Chrome()
        # driver.get(url)
        # for i, ad in enumerate(ads):
        #     print_button = driver.find_element(By.LINK_TEXT, "Распечатать объявление")
            # if print_button:
                # print_button.click()
                # time.sleep(5)
                # save_as_pdf_button = driver.find_element('span', {'class': 'action-name', 'text': 'Сохранить в PDF'})
                # if save_as_pdf_button:
                #     save_as_pdf_button.click()
                #     time.sleep(5)
                #     filename = max([os.path.join(root, name) for root, dirs, files in os.walk(os.path.expanduser('~\\Downloads')) for name in files], key=os.path.getctime)
                #     new_filename = f"{pdf_folder_path}/Siebel-DI_{inn}_{i+1}.pdf"
                #     os.rename(filename, new_filename)
                #     print(f"Файл {new_filename} сохранен")
        return "Найдена информация об ИНН", datetime.datetime.now()
    else:
        return "Не найдена информация", datetime.datetime.now()

# Функция для обработки данных
def process_data(df, pdf_folder_path):
    for index, row in df.iterrows():
        inn = row.iloc[7]
        if inn:
            result, date = search_and_save_pdf(inn, pdf_folder_path)
            df.at[index, "Найдено в Ъ"] = result
            df.at[index, "Дата проверки"] = date

# Функция для запуска обработки данных

def main(file_path, pdf_folder_path):
    # Создаем папку для хранения pdf файлов, если она не существует
    os.makedirs(pdf_folder_path, exist_ok=True)
    df = load_data(file_path)
    process_data(df, pdf_folder_path)
    # Сохраняем измененный DataFrame в Excel файл
    output_file_path = file_path.split('.')[0] + '_output.xlsx'
    df.to_excel(output_file_path, index=False)
    print(f"Данные сохранены в файл {output_file_path}")

main('xlsx/clients.xlsx','pdf')