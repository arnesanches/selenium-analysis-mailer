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

# ======== Automação com Selenium: Envio automatizado de email =========
try:
    # Abrindo o Outlook para iniciar o processo de envio do email
    driver.get("https://outlook.live.com/owa/")

    # Aguarda até que o botão "Entrar" esteja disponível
    wait = WebDriverWait(driver, 30)
    botao_entrar = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Entrar")))
    botao_entrar.click()

    # Pausa para garantir que a página de login carregue completamente
    time.sleep(5)

    # Inicia o login inserindo o endereço do email criado para o código
    pyautogui.write("teste.automatizado@outlook.com")
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(5)
    # Insere senha para concluir o login
    pyautogui.write("TesteAutomatizado123")
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(5)
    # Direciona para a opção "Novo email"
    pyautogui.hotkey('crtl', 'alt')
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(3)
    # Insere o endereço de email do destinatário
    pyautogui.write("diretoria@empresa.com")
    time.sleep(2)
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(2)
    # Informa o assunto da mensagem
    pyautogui.write("Resumo de Vendas - Dezembro")
    pyautogui.press('enter')
    time.sleep(2)
    # Cola a mensagem gerada no corpo do email
    pyperclip.copy(mensagem)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(5)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(2)
    # Envia o email
    pyautogui.hotkey("ctrl", "enter")
    time.sleep(5)

    # Navega com a tecla Tab até acessar a opção "Sair" para realizar o logout
    for i in range(24):
        pyautogui.press('tab')

    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    # Conclui o logout
    pyautogui.press('enter')