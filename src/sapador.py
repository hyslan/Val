"""Módulo de start do SAP"""
import time
import os
import subprocess
from typing import Callable, Literal

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv


def down_sap() -> None:
    """Baixa o .tx do SAP"""
    load_dotenv()
    url: str = os.environ["URL"]
    s: Service = Service(
        'src/chromedriver.exe')
    opt: Options = Options()
    opt.add_argument('--headless=new')
    opt.add_argument('--allow-running-insecure-content')
    opt.add_argument('--ignore-certificate-errors')
    opt.add_argument(
        f'--unsafely-treat-insecure-origin-as-secure={os.environ["URL"]}')
    opt.add_experimental_option('prefs', {
        "download.default_directory": os.getcwd() + "\\src",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True

    })
    driver: WebDriver = webdriver.Chrome(service=s, options=opt)
    # Navegar até a página de login
    driver.get(url)
    wait: WebDriverWait = WebDriverWait(driver, 180)
    wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            '//div[@title="SiiS"]'
        ))
    )
    btn_siis: WebElement = driver.find_element(
        By.XPATH,
        '//div[@title="SiiS"]'
    )
    btn_siis.click()
    wait.until(file_downloaded("tx.sap"))
    print("Arquivo baixado.")
    driver.quit()
    # Caminho para o arquivo "tx.sap"
    caminho_arquivo: str = os.getcwd() + "\\src\\tx.sap"

    # Tenta executar o comando
    try:
        subprocess.run(["powershell", "start", '"' + caminho_arquivo + '"'],
                       shell=True, check=False)
        time.sleep(5)
        # Verifica se o processo está em execução
        if not is_process_running("powershell.exe"):
            print("Erro: O arquivo não foi aberto corretamente.")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao iniciar o arquivo tx.sap: {e}")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")


def file_downloaded(filename: str) -> Callable[[WebDriver], Literal[False] | bool]:
    """Verifica se o arquivo foi baixado completamente"""
    def predicate(driver: WebDriver) -> Literal[False] | bool:
        files: list[str] = os.listdir(os.getcwd() + "\\src")
        return any(file.endswith(filename) for file in files)

    return predicate


def is_process_running(process_name: str):
    """Verifica se o processo está em execução"""
    try:
        subprocess.check_output(
            f'tasklist /FI "IMAGENAME eq {process_name}"', shell=True)
        return True
    except subprocess.CalledProcessError:
        return False
