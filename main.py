from tkinter.filedialog import askopenfilename
import openpyxl

from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
import os

def read_excel():
    # Создаем папки
    # ==============================================
    path = f"C:{os.sep}Program Files{os.sep}WAusers"
    try:
        os.makedirs(path)
    except FileExistsError:
        pass
    # ==============================================

    # Настроиваем driver
    # ==============================================================
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 30)
    url = "www.google.com"
    driver.get(url)
    # ==============================================================

    # Создаем объкт класса для обрабокти данных
    wb = openpyxl.load_workbook(askopenfilename())

    # Получаем активный лист
    sheet = wb.active

    # Начало отчета ячейки,всегда будет начинаться от 2
    i = 2

    while True:
        number_1 = sheet[f'A{i}'].value
        number_2 = sheet[f'B{i}'].value
        text = sheet[f'C{i}'].value

        if number_1 is None:
            return False
        
        elif sheet[f'D{i}'].value is not None:
            return False
        i += 1

def main():
    pass


if __name__ == "__main__":
    read_excel()