"""Módulo de start do SAP."""
import os
import re
import subprocess
import time
from collections.abc import Callable
from typing import TYPE_CHECKING, Literal

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

if TYPE_CHECKING:
    from selenium.webdriver.remote.webelement import WebElement


def down_sap() -> str:
    """Baixa o .tx do SAP."""
    load_dotenv()
    url: str = os.environ["URL"]
    s: Service = Service(
        "chromedriver.exe")
    opt: Options = Options()
    opt.add_argument("--headless=new")
    opt.add_argument("--allow-running-insecure-content")
    opt.add_argument("--ignore-certificate-errors")
    opt.add_argument(
        f'--unsafely-treat-insecure-origin-as-secure={os.environ["URL"]}')
    opt.add_experimental_option("prefs", {
        "download.default_directory": os.getcwd() + "\\shortcut",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,

    })
    driver: WebDriver = webdriver.Chrome(service=s, options=opt)
    # Navegar até a página de login
    driver.get(url)
    wait: WebDriverWait = WebDriverWait(driver, 180)
    wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            '//div[@title="SiiS"]',
        )),
    )
    btn_siis: WebElement = driver.find_element(
        By.XPATH,
        '//div[@title="SiiS"]',
    )
    btn_siis.click()
    wait.until(file_downloaded("tx.sap"))
    driver.quit()
    # Caminho para o arquivo "tx.sap"
    caminho_arquivo: str = os.getcwd() + "\\shortcut\\tx.sap"
    # Get the token SSO
    with open(caminho_arquivo) as f:
        txt = f.read()

    scan = re.search(r'at="MYSAPSSO2=(.*)"', txt)
    token = scan.group(1)
    # Tenta executar o comando
    try:
        subprocess.run(["powershell", "start", '"' + caminho_arquivo + '"'],
                       shell=True, check=False)
        time.sleep(5)
        # Verifica se o processo está em execução
        if not is_process_running("powershell.exe"):
            pass
    except subprocess.CalledProcessError:
        pass
    except Exception:
        pass

    return token


def file_downloaded(filename: str) -> Callable[[WebDriver], Literal[False] | bool]:
    """Verifica se o arquivo foi baixado completamente."""
    def predicate(driver: WebDriver) -> Literal[False] | bool:
        files: list[str] = os.listdir(os.getcwd() + "\\shortcut\\")
        return any(file.endswith(filename) for file in files)

    return predicate


def is_process_running(process_name: str) -> bool | None:
    """Verifica se o processo está em execução."""
    try:
        subprocess.check_output(
            f'tasklist /FI "IMAGENAME eq {process_name}"', shell=True)
        return True
    except subprocess.CalledProcessError:
        return False
