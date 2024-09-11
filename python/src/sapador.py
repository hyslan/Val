"""Módulo de start do SAP."""

import logging
import os
import re
import shlex  # Add this line
import subprocess
import time
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.wait import WebDriverWait

if TYPE_CHECKING:
    from selenium.webdriver.remote.webelement import WebElement

logger = logging.getLogger(__name__)


def down_sap() -> str:
    """Baixa o .tx do SAP."""
    load_dotenv()
    url: str = os.environ["URL"]
    s: Service = Service("chromedriver.exe")
    opt: Options = Options()
    opt.add_argument("--headless=new")
    opt.add_argument("--allow-running-insecure-content")
    opt.add_argument("--ignore-certificate-errors")
    opt.add_argument(f'--unsafely-treat-insecure-origin-as-secure={os.environ["URL"]}')
    opt.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(Path.cwd() / "shortcut"),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
        },
    )
    driver: WebDriver = webdriver.Chrome(service=s, options=opt)
    # Navegar até a página de login
    driver.get(url)
    wait: WebDriverWait[WebDriver] = WebDriverWait(driver, 180)
    wait.until(
        ec.element_to_be_clickable(
            (
                By.XPATH,
                '//div[@title="SiiS"]',
            ),
        ),
    )
    btn_siis: WebElement = driver.find_element(
        By.XPATH,
        '//div[@title="SiiS"]',
    )
    btn_siis.click()
    wait.until(file_downloaded("tx.sap"))
    driver.quit()
    # Caminho para o arquivo "tx.sap"
    caminho_arquivo: str = str(Path.cwd() / "shortcut" / "tx.sap")
    # Get the token SSO
    with Path(caminho_arquivo).open() as f:
        txt = f.read()

    scan = re.search(r'at="MYSAPSSO2=(.*)"', txt)
    token = scan.group(1) if scan is not None else ""
    # Tenta executar o comando
    try:
        subprocess.run(["powershell", "start", shlex.quote(caminho_arquivo)], check=False)
        time.sleep(5)
        # Verifica se o processo está em execução
        if not is_process_running("powershell.exe"):
            pass
    except subprocess.CalledProcessError:
        logger.exception("Erro ao executar o comando de chamada do tx.sap em sapador.py")

    return token


def file_downloaded(filename: str) -> Callable[[WebDriver], bool]:
    """Verifica se o arquivo foi baixado completamente."""

    def predicate(driver: WebDriver) -> bool:  # noqa: ARG001
        files: list[str] = os.listdir(Path.cwd() / "shortcut")
        return any(file.endswith(filename) for file in files)

    return predicate


def is_process_running(process_name: str) -> bool | None:
    """Verifica se o processo está em execução."""
    try:
        subprocess.check_output(["tasklist", "/FI", f"IMAGENAME eq {shlex.quote(process_name)}"])
    except subprocess.CalledProcessError:
        return False
    else:
        return True
