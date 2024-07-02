from selenium import webdriver
from openpyxl import load_workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

class P002_Data_Capture:
    def __init__(self, caminho_arquivo):
        self.caminho = caminho_arquivo
        self.planilha_aberta = load_workbook(filename=self.caminho)
        self.sheet = self.planilha_aberta['Sheet1']

    def T01_setup_browser(self):
        self.navegador = webdriver.Chrome()
        self.espera = WebDriverWait(self.navegador, 10)

    def T02_fill_form(self, nome, email, telefone, sexo, sobre):
        self.navegador.get("https://pt.surveymonkey.com/r/WLXYDX2")

        campo_nome = self.espera.until(EC.presence_of_element_located((By.NAME, "166517069")))
        campo_nome.send_keys(nome)

        campo_email = self.espera.until(EC.presence_of_element_located((By.NAME, "166517072")))
        campo_email.send_keys(email)

        campo_telefone = self.espera.until(EC.presence_of_element_located((By.NAME, "166517070")))
        campo_telefone.send_keys(telefone)

        if sexo == "Fem":
            campo_sexo = self.espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="166517071_1215509813_label"]/span[1]')))
            campo_sexo.click()
        else:
            campo_sexo = self.espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="166517071_1215509812_label"]/span[1]')))
            campo_sexo.click()

        campo_sobre = self.espera.until(EC.presence_of_element_located((By.NAME, "166517073")))
        campo_sobre.send_keys(sobre)

        botao_enviar = self.espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="patas"]/main/article/section/form/div[2]/button')))
        botao_enviar.click()

    def T03_process_data(self):
        for linha in range(2, len(self.sheet['A']) + 1):
            nome = self.sheet[f'A{linha}'].value
            email = self.sheet[f'B{linha}'].value
            telefone = self.sheet[f'C{linha}'].value
            sexo = self.sheet[f'D{linha}'].value
            sobre = self.sheet[f'E{linha}'].value

            self.T01_setup_browser()
            self.T02_fill_form(nome, email, telefone, sexo, sobre)
            self.navegador.quit()

if __name__ == "__main__":
    caminho_arquivo = "C:\\Users\\Pichau\\Documents\\GitHub\\bot_py_02\\output.xlsx"
    capturador = P002_Data_Capture(caminho_arquivo)
    capturador.T03_process_data()
