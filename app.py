from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import *
from selenium.webdriver.support import expected_conditions as EC
from datetime import date
from info import *
from docx import Document
from docx.shared import Inches

def iniciar_driver():
    try:
        chrome_options = Options()

        arguments = [
        '--lang=pt-BR',
        '--window-size=1200,800',
        '--incognito',
        '--disable-infobars'
        '--force-device-scale-factor=0.8'
        ]

        for argument in arguments:
            chrome_options.add_argument(argument)

        chrome_options.add_experimental_option('prefs', {
            'download.prompt_for_download': False,
            'profile.default_content_setting_values.notifications': 2,
            'profile.default_content_setting_values.automatic_downloads': 1,
        })

        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
        
        wait = WebDriverWait(
            driver=driver,
            timeout=10,
            poll_frequency=1,
            ignored_exceptions=[
                NoSuchElementException,
                ElementNotVisibleException,
                ElementNotSelectableException
            ]
        )

        return driver, wait

    except WebDriverException as e:
        print(f'Erro ao iniciar o driver: {e}')
        return None, None

def open_url(url):
    try:
        driver, wait = iniciar_driver()
        driver.get(url)

        return driver, wait
    
    except Exception as e:
        print(f'Erro ao abrir a URL: {e}')
        return None, None

def colect_usd(current_usd_xpath, wait):
    try:
        return round(float(wait.until(EC.visibility_of_all_elements_located((By.XPATH, current_usd_xpath)))[0].text.replace(',', '.')),2)

    except TimeoutException as e:
        print('Erro ao obter a cotação do dolar')

def get_current_date(date):
    return date.today().strftime("%d/%m/%y")

def save_screenshot(driver):
    driver.save_screenshot('cotacao.png')

def write_doc_content(current_usd, current_date, url):
    heading = f'Cotação atual do dolar R${current_usd} - {current_date}'
    content = f'''
    O dólar está no valor de R${current_usd}, na data {current_date}.
    Valor cotado no {url}
    Print da cotação atual:
    '''

    return heading, content

def create_doc(heading, content):
    document = Document()
    document.add_heading(heading, level=0)
    paragrafo = document.add_paragraph(content)
    document.add_picture('./cotacao.png', width=Inches(6), height=Inches(3.5))
    document.save('Cotação atual do dolar.docx')

def main():
    driver, wait = open_url(url)
    current_usd_float = colect_usd(current_usd_xpath, wait)
    current_date = get_current_date(date)
    save_screenshot(driver)
    heading, content = write_doc_content(current_usd_float, current_date, url)
    create_doc(heading, content)

if __name__ == '__main__':
    main()