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

time.sleep(2) # Aguarda para garantir que SAP está aberto e logado


# Cadastro de Fornecedor usando fk01 do SAP
def cadastrar_fornecedor(nomeEmpresa, endRua, endNumb, endCEP, endCidade, endPaís, endEstado, languegePT, telFixo, telCel, email):
    #(nome_empresa, endereco, cidade, estado, telefone, email):
    try:
        # If Not IsObject(application) Then
        # Set SapGuiAuto  = GetObject("SAPGUI")
        sapguiauto = win32com.client.GetObject("SAPGUI")
        #    Set application = SapGuiAuto.GetScriptingEngine
        application = sapguiauto.GetScriptingEngine
        # End If
        # If Not IsObject(connection) Then
        #    Set connection = application.Children(0)
        connection = application.Children(0)
        # End If
        # If Not IsObject(session) Then
        #    Set session    = connection.Children(0)
        session = connection.Children(0)
        # End If

        print(type(session))
        print("Executado com sucesso")


        # session.findById("wnd[0]").resizeWorkingPane 131,33,false
        session.findById("wnd[0]/tbar[0]/okcd").text = "fk01"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA02P01:SAPLBUD0:1130/cmbBUS000FLDS-TITLE_MEDI").key = "0003"
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA02P02:SAPLBUD0:1200/txtBUT000-NAME_ORG1").text = nomeEmpresa
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "txtADDR1_DATA-STREET").text = endRua
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "txtADDR1_DATA-HOUSE_NUM1").text = endNumb
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "txtADDR1_DATA-POST_CODE1").text = endCEP
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "txtADDR1_DATA-CITY1").text = endCidade
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "ctxtADDR1_DATA-COUNTRY").text = endPaís
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "ctxtADDR1_DATA-REGION").text = endEstado
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "cmbADDR1_DATA-LANGU").key = languegePT
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "txtSZA1_D0100-TEL_NUMBER").text = telFixo
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "txtSZA1_D0100-MOB_NUMBER").text = telCel
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "txtSZA1_D0100-SMTP_ADDR").text = email
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "txtSZA1_D0100-SMTP_ADDR").setFocus
        session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                        "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                        "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/"
                        "subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/"
                        "txtSZA1_D0100-SMTP_ADDR").caretPosition = 22
        session.findById("wnd[0]/tbar[0]/btn[11]").press()

    except Exception as e:
        print(f"Erro ao cadastrar fornecedor: {e}")
        return None
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
email = "contato@email.com"
    
if __name__ == '__main__':
    openSAP()
    time.sleep(5)  # Aguarda para garantir que SAP está aberto e logado 
    session = logonSAP(usuario_sap, senha_sap, conexao_sap)
    cadastrar_fornecedor(nomeEmpresa, endRua, endNumb, endCEP, endCidade, endPaís, endEstado, languegePT, telFixo, telCel, email)
    
