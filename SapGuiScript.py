import bot
import os
# Import Pandas paara manipulação de dados Excel
import pandas as pd

# Abre o SAP - cod limpo importando a função openSAP
bot.openSAP()

# Variáveis para login
user = os.getenv("USUARIO_SAP")
password = os.getenv("SENHA_SAP")
connection = "S4H [sapvirtual1.ddns.net]"  # Nome exato da conexão de acesso ao SAP pós aberto

bot.logonSAP(user, password, connection)


class listaFornecedores():

    # Caminho do arquivo Excel
    path = r"E:\RPA\BotCity\Projetos\SAPproject2\resources\fornecedores(1).xlsx"
    
    def lerDados(self, path, range=None, sheet="Sheet1", head=True):

        # Trazer os dados do Excel
        df = pd.read_excel(path, sheet_name=sheet, header=0 if head else None)
        dados = df.values.tolist()  # Converte o DataFrame em uma lista de listas
        return dados
    
    def cadastroFornecedores(self):
        # Aqui o loop na planilha para cadastrar os fornecedores
        # Leros dados do Excel
        dados = self.lerDados(self.path)

        # Percorrer toda planilha e cadastrar os fornecedores
        for linha in dados:
            nomeEmpresa, endRua, endNumb, endCEP, endCidade, endPaís, endEstado, languegePT, telFixo, telCel, email = linha

            # Chamando a Função de cadastro de fornecedor 
            bot.cadastrar_fornecedor(nomeEmpresa, endRua, endNumb, endCEP, endCidade, endPaís, endEstado, languegePT, telFixo, telCel, email)

# Criar a instância e chamar a função
fornecedores = listaFornecedores()
fornecedores.cadastroFornecedores()


