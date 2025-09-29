import win32com.client  # type: ignore
import subprocess
import time
import os
from dotenv import load_dotenv
from datetime import datetime
import psutil

def extracao_sap():
    """Executa a extração da transação KE5Z no SAP e exporta para Excel"""

    load_dotenv()

    USERNAMESAP = os.getenv('USERNAMESAP')
    PASSWORDSAP = os.getenv('PASSWORDSAP')
    PATHSAP = os.getenv('SUBPATHSAP')
    

    # === FUNÇÃO PARA FECHAR TODAS AS INSTÂNCIAS DO SAP ===
    def fechar_sap():
        """Fecha todas as instâncias do SAP Logon e SAP GUI"""
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and 'saplogon.exe' in proc.info['name'].lower():
                    proc.terminate()
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and ('sapgui.exe' in proc.info['name'].lower() or 'sapfront.exe' in proc.info['name'].lower()):
                    proc.terminate()
            time.sleep(2)
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and ('saplogon.exe' in proc.info['name'].lower() or 
                                        'sapgui.exe' in proc.info['name'].lower() or 
                                        'sapfront.exe' in proc.info['name'].lower()):
                    proc.kill()
            time.sleep(1)
            print("Todas as instâncias do SAP foram fechadas.")
        except Exception as e:
            print(f"Erro ao fechar SAP: {e}")

    # === FUNÇÃO PARA FECHAR EXCEL ===
    def fechar_excel():
        """Fecha todas as instâncias do Excel"""
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and 'excel.exe' in proc.info['name'].lower():
                    proc.terminate()
            time.sleep(2)
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and 'excel.exe' in proc.info['name'].lower():
                    proc.kill()
            time.sleep(1)
            print("Todas as instâncias do Excel foram fechadas.")
        except Exception as e:
            print(f"Erro ao fechar Excel: {e}")

    # === FECHA SAP E EXCEL EXISTENTES ANTES DE INICIAR ===
    fechar_sap()
    fechar_excel()
    time.sleep(2)

    # === INICIALIZAÇÃO E CONEXÃO SAP ===
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        for i in range(application.Children.Count):
            try:
                connection = application.Children(i)
                connection.CloseSession()
                time.sleep(1)
            except:
                continue
        subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
        time.sleep(5)
    except:
        subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
        time.sleep(5)

    # Reconecta após fechar tudo
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        time.sleep(3)
        connection = application.OpenConnection("PRD_New", True)
        session = connection.Children(0)

        try:
            for i in range(5):
                try:
                    session.findById(f"wnd[{i}]").sendVKey(0)
                    time.sleep(1)
                except:
                    break
        except:
            pass

    except Exception as e:
        print(f"Erro na conexão SAP: {e}")
        return

    # === LOGIN ===
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "100"
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = f"{USERNAMESAP}"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = f"{PASSWORDSAP}"
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)

    # === TRATA TELA DE LOGON MÚLTIPLO ===
    try:
        if session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2"):
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(1)
    except:
        pass

    # === TRANSAÇÃO KE5Z ===
    session.findById("wnd[0]/tbar[0]/okcd").Text = "KE5Z"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(3)

    # === SELECIONAR VARIANTE FATURAMENTO ===
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    time.sleep(2)

    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell(5, "TEXT")
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "5"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    time.sleep(2)

    # === INSERIR PERÍODO (MÊS ATUAL) ===
    mes_atual = datetime.now().strftime("%m")
    session.findById("wnd[0]/usr/ctxtPOPER-LOW").text = mes_atual
    time.sleep(1)
    session.findById("wnd[0]/usr/ctxtPOPER-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtPOPER-LOW").caretPosition = 2
    time.sleep(1)

    # === EXECUTAR (F8) ===
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(3)

    # === EXPORTAR PARA EXCEL (SUBSTITUINDO SE JÁ EXISTIR) ===
    arquivo_excel = os.path.join(PATHSAP, "externosap.xlsx")

    if os.path.exists(arquivo_excel):
        os.remove(arquivo_excel)
        print("Arquivo existente removido para sobrescrever.")

    session.findById("wnd[0]").sendVKey(16)
    time.sleep(2)

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = rf"{PATHSAP}"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "externosap.xlsx"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
    time.sleep(1)

    session.findById("wnd[1]/tbar[0]/btn[11]").press()  # Salvar
    time.sleep(2)

    print("Exportação concluída com sucesso!")
    
    # === FECHAR EXCEL APÓS SALVAR ===
    print("Fechando o Excel...")
    fechar_excel()
    time.sleep(1)
    
    # === FECHAR SAP APÓS CONCLUSÃO ===
    print("Fechando o SAP...")
    fechar_sap()
    
    print("Processo finalizado completamente!")

