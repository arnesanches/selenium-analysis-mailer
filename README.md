### English

# Selenium Analysis Mailer

This repository contains a script that performs data analysis on an Excel file, generates insights, and automatically sends an email with this information via Outlook. The script also logs out and closes the browser after completing the email sending process. The script is available in the following formats:

Python File: script.py
Jupyter Notebook: script.ipynb

#### Features
Data Analysis:

The script processes data from the Excel file to identify:

The best-selling product and its total quantity
.
The store with the highest revenue and its corresponding value.

Automated Email Sending:

Uses Selenium to access Outlook on the web and automatically send an email containing the generated insights.

Logout and Browser Closure:

After sending the email, the script logs out of Outlook and closes the browser to ensure security and clean up the environment.

#### Important Notes

Security:

Email credentials are embedded in the code as plain text. For a production environment, it is recommended to use a secret manager or environment variables to protect sensitive information.
Outlook Layout Dependency:

The Outlook Online layout changes frequently. If the script stops working correctly, it might be necessary to adjust the selectors or commands used to interact with the interface.

#### Requirements

Ensure that the Excel file Vendas - Dez.xlsx is located in the same folder as the script. This file is essential for executing the code as it serves as the data source for the analysis.
Install the Google Chrome browser on your machine, as it is used by Selenium for automation.
Install the required libraries before running the code. To install them, use the following command:

pip install pandas selenium webdriver-manager pyautogui pyperclip

#### Script Execution

.py Format:

Run the script directly using the command: python script.py.

.ipynb Format:

Open the notebook in a Jupyter environment and execute the cells sequentially.

#### Reason for the Two Formats

The script is available in both .py and .ipynb formats to cater to different needs:

.py Format: Ideal for direct execution in production environments or integration into automated pipelines.

.ipynb Format: Suitable for study, demonstrations, or experimentation in an interactive environment such as Jupyter Notebook.

Contact

If you encounter any issues or have suggestions, feel free to open an issue in this repository!

---

