from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys

import pandas as pd

# Install the version of the chrome web driver in the machine.
service = Service(ChromeDriverManager().install())
browser = webdriver.Chrome(service=service) # Instance of new page in chrome

def reply(cpf, email, description, value):
  browser.get("https://forms.gle/4N5s1BXLQCdFFJ896") # Opening the google form.

  # CPF Input
  cpfInput = browser.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input')
  cpfInput.send_keys(cpf)

  # Email Input
  emailInput = browser.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input')
  emailInput.send_keys(email)

  # Description Input
  descriptionInput = browser.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input')
  descriptionInput.send_keys(description)

  # Values Input
  valueInput = browser.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input')
  valueInput.send_keys(value)

  # Send the form 
  sendBtn = browser.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div')
  sendBtn.send_keys(Keys.RETURN) # Pressing Enter to send


table = pd.read_excel("Emit.xlsx")

for i, cpf in enumerate(table["CPF"]):
  email = table.loc[i, "Email"]
  description = table.loc[i, "Descrição"]
  value = table.loc[i, "Valor"]

  reply(cpf, email, description, str(value))