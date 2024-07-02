from selenium import webdriver
from openpyxl import load_workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By\

caminho = "Dados.xlsx"
planilha_aberta = load_workbook(filename=caminho)
sheet = planilha_aberta['Sheet1']

for linha in range(2, len(sheet['A'] + 1)):
    nome = sheet[f'A{linha}'].value
    email = sheet[f'B{linha}'].value
    telefone = sheet[f'C{linha}'].value
    sexo = sheet[f'D{linha}'].value
    sobre = sheet[f'E{linha}'].value

    navegador = webdriver.Chrome()
    navegador.get("https://pt.surveymonkey.com/r/WLXYDX2")

    espera = WebDriverWait(navegador, 10)\

    campo_nome = espera.until(expected_conditions.presence_of_element_located((By.NAME, "166517069")))
    campo_nome.send_keys(nome)

    campo_email = espera.until(expected_conditions.presence_of_element_located((By.NAME, "166517072")))
    campo_email.send_keys(email)

    campo_telefone = espera.until(expected_conditions.presence_of_element_located((By.NAME, "166517070")))
    campo_telefone.send_keys(telefone)

    if sexo == "Fem":
        campo_sexo = espera.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="166517071_1215509813_label"]/span[1]')))
        campo_sexo.click()
    else:
        campo_sexo = espera.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="166517071_1215509812_label"]/span[1]')))
        campo_sexo.click()

    campo_sobre = espera.until(expected_conditions.presence_of_element_located((By.NAME, "166517073")))
    campo_sobre.send_keys(sobre)

    botao_enviar = espera.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="patas"]/main/article/section/form/div[2]/button')))
    botao_enviar.click()