### Project Demonstration / Demonstração do Projeto

![Project Demonstration / Demonstração do Projeto](https://github.com/arnesanches/selenium-analysis-mailer/blob/main/Anima%C3%A7%C3%A3o.gif?raw=true)

#### English

# Selenium Analysis Mailer

This Python script combines the Selenium, Pandas, and PyAutoGUI libraries to automate the analysis of data from an Excel file and send a report with the results of this analysis via email through Outlook, optimizing workflow and saving time. The complete automation, including logout and browser closure, ensures the security and organization of the process. The script is available in the following formats:

* Python File: script.py

* Jupyter Notebook: script.ipynb

## Features

### Data Analysis:

Using the Pandas library, the script processes data from the Excel file to identify:

* The best-selling product and its total quantity.

* The store with the highest revenue and its corresponding value.

### Automated Email Sending:

* It uses the Selenium and PyAutoGUI libraries to access Outlook on the web and automatically send an email containing the generated insights.

### Logout and Browser Closure:

* After sending the email, the script logs out of Outlook and closes the browser to ensure security and clean up the environment.

## Important Notes

### Security:

Email credentials are embedded in the code as plain text. For a production environment, it is recommended to use a secret manager or environment variables to protect sensitive information.

### Outlook Layout Dependency:

The Outlook Online layout changes frequently. If the script stops working correctly, it might be necessary to adjust the selectors or commands used to interact with the interface.

## Requirements

* Install Python 3.7 or higher. You can download the latest version at [https://www.python.org/downloads/](https://www.python.org/downloads/).

* Ensure that the Excel file Vendas - Dez.xlsx is located in the same folder as the script. This file is essential for executing the code as it serves as the data source for the analysis.

* Install the Google Chrome browser on your machine, as it is used by Selenium for automation.

* Install the required libraries before running the code. To install them, use the following command:

#### pip install pandas selenium webdriver-manager pyautogui pyperclip

## Script Execution

### .py Format:

* Run the script by clicking the 'Run' button or using the command:

  '''bash
  python script.py
  '''

### .ipynb Format:

* Open the notebook in a Jupyter environment and execute the cells sequentially.

If you wish to interrupt the automation before it is completed, position the mouse pointer in the top left corner of the screen and wait for about five seconds.

## Reason for the Two Formats

The script is available in both .py and .ipynb formats to cater to different needs:

* .py Format: Ideal for direct execution in production environments or integration into automated pipelines.

* .ipynb Format: Suitable for study, demonstrations, or experimentation in an interactive environment such as Jupyter Notebook.

## Contributions

Contributions are welcome! Feel free to open issues or pull requests in this repository.

## License

This project is licensed under the MIT License.

---

#### Português 

# Selenium Analysis Mailer

Este script Python combina as bibliotecas Selenium, Pandas e PyAutoGUI para automatizar a análise de dados de um arquivo Excel e o envio de um relatório com os resultados dessa análise por e-mail através do Outlook, otimizando o fluxo de trabalho e economizando tempo. A automação completa, incluindo logout e fechamento do navegador, assegura a segurança e a organização do processo. O script está disponível nos seguintes formatos:

* Arquivo Python: script.py

* Notebook Jupyter: script.ipynb

## Funcionalidades

### Análise de Dados:

Utilizando a biblioteca Pandas, o script processa os dados do arquivo Excel para identificar:

* O produto mais vendido e sua quantidade total.

* A loja com maior faturamento e o valor correspondente.

### Envio Automatizado de E-mail:

* Utiliza as bibliotecas Selenium e PyAutoGUI para acessar o Outlook na web e enviar automaticamente um e-mail contendo os insights gerados.

### Logout e Fechamento do Navegador:

* Após o envio do e-mail, o script realiza logout do Outlook e fecha o navegador para garantir segurança e limpeza do ambiente.

## Observações Importantes

### Segurança:

As credenciais do e-mail estão embutidas no código como texto plano. Para um ambiente de produção, recomenda-se o uso de um gerenciador de segredos ou variáveis de ambiente para proteger informações sensíveis.

### Dependência do Layout do Outlook:

O layout do Outlook Online muda com frequência. Caso o script pare de funcionar corretamente, pode ser necessário ajustar os seletores ou comandos usados para interagir com a interface.

## Requisitos

* Instale o Python 3.7 ou superior. Você pode baixar a versão mais recente em [https://www.python.org/downloads/](https://www.python.org/downloads/).

* Certifique-se de que o arquivo Excel Vendas - Dez.xlsx esteja localizado na mesma pasta que o script. Este arquivo é essencial para a execução do código, pois serve como fonte de dados para a análise.

* Instale o navegador Google Chrome na máquina, pois ele é usado pelo Selenium para automação.

* Instale as bibliotecas necessárias antes de executar o código. Para instalá-las, use o seguinte comando:

#### pip install pandas selenium webdriver-manager pyautogui pyperclip

## Execução do Script

### Formato .py:

* Execute o script clicando no botão 'Run' ou através do comando:

'''bash
  python script.py
  '''
  
### Formato .ipynb:

* Abra o notebook em um ambiente Jupyter e execute as células sequencialmente.

Caso queira interromper a automação antes de ser concluída, direcione a seta do mouse para o canto superior esquerdo da tela e aguarde cerca de cinco segundos.

## Motivo para os Dois Formatos

O script está disponível tanto no formato .py quanto no formato .ipynb para atender a diferentes necessidades:

* Formato .py: Ideal para execução direta em ambientes de produção ou integração em pipelines automatizados.

* Formato .ipynb: Adequado para estudos, demonstrações ou experimentações em um ambiente interativo, como o Jupyter Notebook.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests neste repositório.

## Licença

Este projeto está licenciado sob a MIT License.


