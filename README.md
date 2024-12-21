### English

# Selenium Analysis Mailer

### Description

This repository contains a script that performs data analysis on an Excel file, generates insights, and automatically sends an email with this information via Outlook. The script also logs out and closes the browser after completing the email sending process. The script is available in the following formats:

#### Python File: script.py

#### Jupyter Notebook: script.ipynb

### Features

#### Data Analysis:

The script processes data from the Excel file to identify:

* The best-selling product and its total quantity.

* The store with the highest revenue and its corresponding value.

#### Automated Email Sending:

Uses Selenium to access Outlook on the web and automatically send an email containing the generated insights.

#### Logout and Browser Closure:

After sending the email, the script logs out of Outlook and closes the browser to ensure security and clean up the environment.

### Important Notes

#### Security:

Email credentials are embedded in the code as plain text. For a production environment, it is recommended to use a secret manager or environment variables to protect sensitive information.

#### Outlook Layout Dependency:

The Outlook Online layout changes frequently. If the script stops working correctly, it might be necessary to adjust the selectors or commands used to interact with the interface.

### Requirements

* Ensure that the Excel file Vendas - Dez.xlsx is located in the same folder as the script. This file is essential for executing the code as it serves as the data source for the analysis.

* Install the Google Chrome browser on your machine, as it is used by Selenium for automation.

* Install the required libraries before running the code. To install them, use the following command:

#### pip install pandas selenium webdriver-manager pyautogui pyperclip

### Script Execution

#### .py Format:

Run the script directly using the command: python script.py.

#### .ipynb Format:

Open the notebook in a Jupyter environment and execute the cells sequentially.

### Reason for the Two Formats

The script is available in both .py and .ipynb formats to cater to different needs:

#### .py Format: Ideal for direct execution in production environments or integration into automated pipelines.

#### .ipynb Format: Suitable for study, demonstrations, or experimentation in an interactive environment such as Jupyter Notebook.

### Contact

If you encounter any issues or have suggestions, feel free to open an issue in this repository!

---

### Português 

# Selenium Analysis Mailer

### Descrição

Este repositório contém um script que realiza a análise de dados de um arquivo Excel, gera insights e envia automaticamente um e-mail com essas informações através do Outlook. O script também realiza logout e fecha o navegador após concluir o envio do e-mail. O script está disponível nos seguintes formatos:

#### Arquivo Python: script.py

#### Notebook Jupyter: script.ipynb

### Funcionalidades

#### Análise de Dados:

O script processa os dados do arquivo Excel para identificar:

* O produto mais vendido e sua quantidade total.

* A loja com maior faturamento e o valor correspondente.

#### Envio Automatizado de E-mail:

Utiliza o Selenium para acessar o Outlook na web e enviar automaticamente um e-mail contendo os insights gerados.

#### Logout e Fechamento do Navegador:

Após o envio do e-mail, o script realiza logout do Outlook e fecha o navegador para garantir segurança e limpeza do ambiente.

### Observações Importantes

#### Segurança:

As credenciais do e-mail estão embutidas no código como texto plano. Para um ambiente de produção, recomenda-se o uso de um gerenciador de segredos ou variáveis de ambiente para proteger informações sensíveis.

#### Dependência do Layout do Outlook:

O layout do Outlook Online muda com frequência. Caso o script pare de funcionar corretamente, pode ser necessário ajustar os seletores ou comandos usados para interagir com a interface.

### Requisitos

* Certifique-se de que o arquivo Excel Vendas - Dez.xlsx esteja localizado na mesma pasta que o script. Este arquivo é essencial para a execução do código, pois serve como fonte de dados para a análise.

* Instale o navegador Google Chrome na máquina, pois ele é usado pelo Selenium para automação.

* Instale as bibliotecas necessárias antes de executar o código. Para instalá-las, use o seguinte comando:

#### pip install pandas selenium webdriver-manager pyautogui pyperclip

### Execução do Script

#### Formato .py:

Execute o script diretamente com o comando: python script.py.

#### Formato .ipynb:

Abra o notebook em um ambiente Jupyter e execute as células sequencialmente.

### Motivo para os Dois Formatos

O script está disponível tanto no formato .py quanto no formato .ipynb para atender a diferentes necessidades:

#### Formato .py: Ideal para execução direta em ambientes de produção ou integração em pipelines automatizados.

#### Formato .ipynb: Adequado para estudos, demonstrações ou experimentações em um ambiente interativo, como o Jupyter Notebook.

### Contato

Se encontrar problemas ou tiver sugestões, sinta-se à vontade para abrir uma issue neste repositório!

