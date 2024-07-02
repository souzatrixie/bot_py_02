import xlsxwriter
import os

class P001_Create_Data:
    def __init__(self, caminho_arquivo):
        self.caminho_arquivo = caminho_arquivo
        self.workbook = xlsxwriter.Workbook(caminho_arquivo)
        self.sheet = self.workbook.add_worksheet()

    def T01_add_headers(self):
        self.sheet.write("A1", "nome")
        self.sheet.write("B1", "email")
        self.sheet.write("C1", "telefone")
        self.sheet.write("D1", "sexo")
        self.sheet.write("E1", "sobre")

    def T02_add_data(self):
        nomes = ["Beatriz Souza", "Elton Aquino", "Adriana M. Souza", "Evie Maria", "Isabela Gouvea"]
        emails = ["beatriz.souza@gmail.com", "Elton_aq@gmail.com", "adriana.souza@gmail.com", "Eviemaria@gmail.com", "Isabela.GO@gmail.com"]
        telefones = ["(15)97781-8877", "(15)97781-4477", "(15)97781-6677", "(15)95581-8877", "(15)97797-8877"]
        sexos = ["Fem", "Masc", "Fem", "Fem", "Fem"]
        sobres = ["RPA Dev", "Designer", "Assistente Social", "Gato", "Engenheira civil"]

        for i in range(len(nomes)):
            self.sheet.write(i + 1, 0, nomes[i])
            self.sheet.write(i + 1, 1, emails[i])
            self.sheet.write(i + 1, 2, telefones[i])
            self.sheet.write(i + 1, 3, sexos[i])
            self.sheet.write(i + 1, 4, sobres[i])

    def T03_save_and_open(self):
        self.workbook.close()
        os.startfile(self.caminho_arquivo)

if __name__ == "__main__":
    caminho_arquivo = 'C:\\Users\\Pichau\\Documents\\GitHub\\bot_py_02\\output.xlsx'
    creator = P001_Create_Data(caminho_arquivo)
    creator.T01_add_headers()
    creator.T02_add_data()
    creator.T03_save_and_open()