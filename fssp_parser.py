from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from python_anticaptcha import AnticaptchaClient, ImageToTextTask
from selenium.common.exceptions import NoSuchElementException
import win32com.client
import time
import base64
import os

# пути до файлов .xlsx и chromedriver
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_PATH = os.path.join(ROOT_DIR, 'данные физ.лица.xlsx')
OUTPUT_PATH = os.path.join(ROOT_DIR, 'данные по ИП физ.лица.xlsx')
CHROMEDRIVER_PATH = os.path.join(ROOT_DIR, 'chromedriver.exe')

# настройки браузера
options = Options()
options.add_argument('start-maximized')
DRIVER = webdriver.Chrome(
    chrome_options=options,
    executable_path=CHROMEDRIVER_PATH
)

# ФИО и год рождения физ.лица
# заполняется автоматически из функции get_values_from_excel
DEBTOR_BIO = []


# чтение данных физ.лица из файла .xlsx
def get_values_from_excel():
    excel = win32com.client.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(INPUT_PATH)
    sheet = wb.ActiveSheet
    first_name = sheet.Cells(2, 2).value
    DEBTOR_BIO.append(first_name)
    last_name = sheet.Cells(2, 1).value
    DEBTOR_BIO.append(last_name)
    patronymic = sheet.Cells(2, 3).value
    DEBTOR_BIO.append(patronymic)
    birth_date = sheet.Cells(2, 4).value.strftime("%d.%m.%Y")
    DEBTOR_BIO.append(birth_date)
    wb.Close()
    excel.Quit()


# отправка результатов проверки в .xlsx-файл
def send_values_to_excel(debtor, process, document,
                         subject, fssp, officer, count):
    excel = win32com.client.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(OUTPUT_PATH)
    sheet = wb.ActiveSheet
    sheet.Cells(2 + count, 1).value = debtor
    sheet.Cells(2 + count, 2).value = process
    sheet.Cells(2 + count, 3).value = document
    sheet.Cells(2 + count, 4).value = subject
    sheet.Cells(2 + count, 5).value = fssp
    sheet.Cells(2 + count, 6).value = officer
    wb.Save()
    wb.Close()
    excel.Quit()


# заполнение формы поиска по данным физ.лица
def search_debtor(first_name, last_name, patronymic, birth_date):
    close_banner = DRIVER.find_element_by_class_name('tingle-modal__close')
    close_banner.click()
    search_btn = DRIVER.find_element_by_xpath(
        '//*[@id="app"]/main/section[1]/div/div/div/div/div[2]/div/div/div/form/div[8]/div[1]/div/div[1]/a'
    )
    search_btn.click()
    first_name_input = DRIVER.find_element_by_name('is[first_name]')
    last_name_input = DRIVER.find_element_by_name('is[last_name]')
    patronymic_input = DRIVER.find_element_by_name('is[patronymic]')
    birth_date_input = DRIVER.find_element_by_name('is[date]')
    fields = [
        first_name_input,
        last_name_input,
        patronymic_input,
        birth_date_input
    ]
    for field in fields:
        field.clear()
    first_name_input.send_keys(first_name)
    last_name_input.send_keys(last_name)
    birth_date_input.send_keys(birth_date)
    patronymic_input.send_keys(patronymic)
    button = DRIVER.find_element_by_class_name('btn-primary')
    button.click()


# запись каптчи в файл
def download_captcha_image(captcha):
    captcha_base64 = captcha.split(',')[1]
    img_bytes = base64.urlsafe_b64decode(
        captcha_base64 + '=' * (4 - len(captcha_base64) % 4)
    )
    fname = 'captcha.png'
    with open(fname, 'wb') as f:
        f.write(img_bytes)


# обход каптчи с помощью Anticaptcha
def breaking_captcha(img):
    api_key = '22200321de09662a544b1a98679dbd98'
    client = AnticaptchaClient(api_key)
    img_file = open('captcha.png', 'rb')
    task = ImageToTextTask(img_file)
    job = client.createTask(task)
    job.join()
    solution = job.get_captcha_text()
    return solution


# заполнение формы каптчи валидными данными
def send_captcha_solution(captcha_solution):
    captcha_input = DRIVER.find_element_by_name('code')
    try:
        captcha_input.send_keys(captcha_solution)
        time.sleep(5)
        submit_btn = DRIVER.find_element_by_xpath(
            '/html/body/div[5]/div[1]/div[2]/div/div/div/form/div[2]/input[2]'
        )
        submit_btn.click()
    except NoSuchElementException:
        captcha_input.clear()
        send_captcha_solution(captcha_solution)


# поиск ифнормации о физ.лице
def get_debtor_info():
    try:
        rows = DRIVER.find_elements_by_tag_name('tr')[2::]
        count = 0  # подсчет кол-ва непогашенных ИП
        for row in rows:
            cols = row.find_elements_by_tag_name('td')
            if cols[3].text == '':
                debtor = cols[0].text
                process = cols[1].text
                document = cols[2].text
                subject = cols[4].text
                fssp = cols[6].text
                officer = cols[7].text
                send_values_to_excel(debtor, process, document,
                                     subject, fssp, officer, count)
                count += 1
    except NoSuchElementException:
        print('Незавершенных исполнительных производств не найдено')
    finally:
        DRIVER.close()


def main():
    get_values_from_excel()
    DRIVER.get('https://fssprus.ru/')
    time.sleep(5)
    search_debtor(DEBTOR_BIO[0], DEBTOR_BIO[1], DEBTOR_BIO[2], DEBTOR_BIO[3])
    time.sleep(5)
    captcha = DRIVER.find_element_by_id('capchaVisual').get_attribute('src')
    img_bytes = download_captcha_image(captcha)
    captcha_solution = breaking_captcha(img_bytes)
    time.sleep(5)
    send_captcha_solution(captcha_solution)
    time.sleep(5)
    get_debtor_info()


if __name__ == '__main__':
    main()
