import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui
import pyperclip
import time

# ======== Configuração do Selenium: Inicializa o WebDriver para interagir com o navegador =========
chrome_options = Options()
chrome_options.add_argument("--start-maximized")

# Configurando o WebDriver com o ChromeDriverManager e iniciando o navegador
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)