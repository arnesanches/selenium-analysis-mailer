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

# ======== Leitura e Análise do Arquivo Excel: Extração de insights =========
file_path = "Vendas - Dez.xlsx"  # Nome do arquivo na mesma pasta do script

# Carregar o arquivo Excel
data = pd.read_excel(file_path)

# Insight 1: Identifica o produto mais vendido e sua quantidade total
produto_mais_vendido = data.groupby('Produto')["Quantidade"].sum().idxmax()
quantidade_vendida = data.groupby('Produto')["Quantidade"].sum().max()

# Insight 2: Determina a loja com maior faturamento e o valor correspondente
loja_maior_faturamento = data.groupby('ID Loja')["Valor Final"].sum().idxmax()
faturamento_loja = data.groupby('ID Loja')["Valor Final"].sum().max()

# Construir a mensagem com os insights
mensagem = (
    f"Segue o resumo de vendas do mês:\n\n"
    f"- Produto mais vendido: {produto_mais_vendido} (Quantidade: {quantidade_vendida})\n"
    f"- Loja com maior faturamento: {loja_maior_faturamento} (Faturamento: R$ {faturamento_loja:,.2f})\n\n"
    f"Atenciosamente,\nSeu sistema automatizado."
)