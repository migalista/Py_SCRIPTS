import win32com.client
import time

# Conectar ao SAP GUI
sap_gui = win32com.client.GetObject("SAPGUI")
application = sap_gui.GetScriptingEngine
connection = application.Children(0)  # Pode ser diferente se você tiver mais de uma conexão aberta
session = connection.Children(0)      # Primeira sessão ativa

# Agora começamos a simular ações no SAP
session.StartTransaction(Transaction="VA03")  # Acessar transação VA03
time.sleep(1)

# Digitar um número de pedido de vendas (exemplo fictício)
session.FindById("wnd[0]/usr/ctxtVBAK-VBELN").Text = "1234567890"

# Pressionar Enter
session.FindById("wnd[0]").sendVKey(0)

print("Script finalizado com sucesso.")
# Esperar um pouco para garantir que a transação carregue
time.sleep(2)
