# Import for the Desktop Bot
from botcity.core import DesktopBot, Backend
import time
import win32com.client
import os
from dotenv import load_dotenv


app_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

# Configuração do login
usuario_sap = os.getenv("USUARIO_SAP")
senha_sap = os.getenv("SENHA_SAP")
conexao_sap = "S4H [sapvirtual1.ddns.net]"  # Nome exato da conexão


def openSAP():
    # Inicializando o bot 
    bot = DesktopBot()

    # Executa e abre o aplicativo
    bot.execute(app_path)
    time.sleep(3)

    try:
        # Conectando -se ao aplicativo usando os seletores de 'path' e 'title'.
        bot.connect_to_app(Backend.WIN_32, path=app_path, title="SAP Logon 770")
        
    except Exception as e:
        print(f"Erro ao abrir o SAP: {e}")
        return
    time.sleep(3)

def logonSAP(usuario, senha, conexao):
    """Faz login no SAP"""
    try:
        # Aguarda o SAP estar ok
        time.sleep(2)

        # Instância do SAP GUI
        sapguiauto = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine

        # Conectar à conexão do SAP Logon
        connection = application.OpenConnection(conexao, True)
        time.sleep(2)  

        session = connection.Children(0)

        # Login no SAP
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "500"  # Mandante
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = senha
        session.findById("wnd[0]").sendVKey(0)  # Pressiona Enter

        print("Login realizado com sucesso!")

        # Caso abra um popup, aqui é para fechar e continuar
        bot = DesktopBot()
        time.sleep(3)  # Espera para ver se o popup aparece

        if bot.find_element("123", matching=0.97, waiting_time=2000):  # Verifica se o botão existe
            print("Popup detectado! Fechando...")
            bot.click()  # Clica no botão OK do popup

        print("SAP pronto para uso.")

        # Aqui irá manter o SAP aberto
        while True:
            time.sleep(1)  # Mantém o script rodando

    except Exception as e:
        print(f"Erro ao logar no SAP: {e}")
        return None

if __name__ == '__main__':
    openSAP()
    time.sleep(5)  # Aguarda para garantir que SAP está aberto e logado 
    logonSAP(usuario_sap, senha_sap, conexao_sap)