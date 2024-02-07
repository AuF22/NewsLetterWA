from tkinter.filedialog import askopenfilename
import openpyxl
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time


def handler_number(number: str)-> str:
    """Приводит номера телефонов в единный формат для отправки"""
    text = str(number).replace(" ", "")
    text = text.replace("+", "")
    text = f"996{text[-9:]}"
    return(text)


def read_excel():
    # # Создаем папки
    # # ==============================================
    # path = f"C:{os.sep}Program Files{os.sep}WAusers"
    # try:
    #     os.makedirs(path)
    # except FileExistsError:
    #     pass
    # # ==============================================

    # Настроиваем driver
    # ==============================================================
    options = webdriver.ChromeOptions()
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 YaBrowser/23.9.0.2325 Yowser/2.5 Safari/537.36")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)
    url = "https://web.whatsapp.com/"
    driver.get(url)
    time.sleep(30)
    # ==============================================================

    text = ""
    # Подготавливаем ссылку для отправки sms
    url_sms = "https://web.whatsapp.com/send?phone={number}&text={text}"

    # Создаем объект класса для обрабокти данных
    wb = openpyxl.load_workbook(askopenfilename())

    # Получаем активный лист
    sheet = wb.active

    # Начало отчета ячейки,всегда будет начинаться от 2
    i = 2

    # Запускаем программу отправку с бесконечным циклом
    # =================================================
    try:
        count = 0
        while True:
            number_1 = sheet[f'A{i}'].value
            number_2 = sheet[f'B{i}'].value
            text = sheet[f'C{i}'].value

            if number_1 is None:
                return False
            
            elif sheet[f'D{i}'].value is not None:
                return False
            
            text = text.replace(' ', '%20')
            text = text.replace('\n', '%0D%0A')

            if sheet[f'D{i}'].value == 'Отправлено':
                i += 1
                print('Продолжаем')
                continue
            i += 1
            
            try:
                count += 1
                url_phone = url_sms.format(number=handler_number(number_1), text=text)
                driver.get(url_phone)
                
                wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[2]/button')))
                driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').click()
                time.sleep(5)
                print(f"Отправлено: {i=}, {count=}")

                sheet[f'D{i-1}'] = 'Отправлено'
                wb.save('Отправлено.xlsx')
                
            except:
                pass

    except:
        wb.save('Отправлено.xlsx')
    # =================================================================
    

def main():
    pass


if __name__ == "__main__":
    read_excel()