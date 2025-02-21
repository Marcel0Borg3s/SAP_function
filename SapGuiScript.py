import bot
import os

# Abre o SAP - cod limpo importando a função openSAP
bot.openSAP()

# Variáveis para login
user = os.getenv("USUARIO_SAP")
password = os.getenv("SENHA_SAP")
connection = "S4H [sapvirtual1.ddns.net]"  # Nome exato da conexão de acesso ao SAP pós aberto

bot.logonSAP(user, password, connection)

# Chamando a Função de cadastro de fornecedor e criando variáveis para os dados
nomeEmpresa = "TESTE Empresa v10"
endRua = "endRua"
endNumb = "999" 
endCEP = "13500-123"
endCidade = "endCidade"
endPaís = "BR"
endEstado = "SP"
languegePT = "PT"
telFixo = "1932569999"
telCel = "1989891111"
email = "contato@e-mail.com"
bot.cadastrar_fornecedor(nomeEmpresa, endRua, endNumb, endCEP, endCidade, endPaís, endEstado, languegePT, telFixo, telCel, email)






