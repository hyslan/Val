# testador.py
'''Área de testes.'''


# import subprocess

# # Caminho para o executável do SAP GUI
# caminho_sap_gui = "C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe"

# # Caminho para o arquivo "tx.sap"
# caminho_arquivo = "C:\\Users\\irgpapais\\Downloads\\tx.sap"

# # Nome do sistema SAP e cliente a serem extraídos do arquivo "tx.sap"
# nome_sistema_sap = "EP0"
# cliente_sap = "100"
# gui_parm = "/M/erpprdci.ti.sabesp.com.br/S/3908/G/PRODUCAO/UPDOWNLOAD_CP=1160"
# at = "MYSAPSSO2=AjExMDAgAA1wb3J0YWw6MTE3NjE1iAATYmFzaWNhdXRoZW50aWNhdGlvbgEABjExNzYxNQIAAzAwMAMAA1BQMQQADDIwMjMxMDAzMTExNgUABAAAAAgKAAYxMTc2MTX/AQYwggECBgkqhkiG9w0BBwKggfQwgfECAQExCzAJBgUrDgMCGgUAMAsGCSqGSIb3DQEHATGB0TCBzgIBATAiMB0xDDAKBgNVBAMTA1BQMTENMAsGA1UECxMESjJFRQIBADAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjMxMDAzMTExNjQzWjAjBgkqhkiG9w0BCQQxFgQUsk1Z3s2WsOCfDAFIrlRQCWox+mUwCQYHKoZIzjgEAwQwMC4CFQDr09OkkEzRQMFFTT/0oNBcJ42p+QIVAPfSk3yagFQE8IAJEiGe9jrcort7"

# # Comando para abrir o arquivo SAP usando o SAP GUI
# comando = [
#     caminho_sap_gui,
#     f"-system={nome_sistema_sap}",
#     f"-client={cliente_sap}",
#     "-file",
#     caminho_arquivo,
#     f"-guiParm={gui_parm}"  # Correção: movido para dentro da lista

# ]

# # Tenta executar o comando
# try:
#     subprocess.run(comando, check=True)
# except subprocess.CalledProcessError as e:
#     print(f"Erro ao executar o SAP GUI: {e}")
# except Exception as e:
#     print(f"Ocorreu um erro: {e}")
import subprocess

# Caminho para o arquivo "tx.sap"
caminho_arquivo = "C:\\Users\\irgpapais\\Downloads\\tx.sap"

# Tenta executar o comando
try:
    subprocess.run(["powershell", "start", caminho_arquivo], shell=True)
except subprocess.CalledProcessError as e:
    print(f"Erro ao iniciar o arquivo tx.sap: {e}")
except Exception as e:
    print(f"Ocorreu um erro: {e}")
