import os.path
import pandas as pd
from datetime import datetime
from pyzabbix import ZabbixAPI

# Conectar-se ao servidor Zabbix
zabbix_url = 'https://zabbix-client01.compwire.com.br/zabbix.php?action=dashboard.view'
zabbix_user = 'luan.siqueira'
zabbix_password = 'caLu*1128'
zapi = ZabbixAPI(zabbix_url)
zapi.login(zabbix_user, zabbix_password)

# Definir o período de tempo para recuperar os dados
time_from = int(datetime(2024, 6, 1, 0, 0, 0).timestamp())
time_till = int(datetime(2024, 6, 3, 23, 59, 59).timestamp())

# Recuperar dados do Zabbix (substitua isso com sua própria lógica)
# Exemplo: obter os 10 itens principais do Zabbix
items = zapi.item.get(output=['name', 'lastvalue'], sortfield='lastvalue', sortorder='DESC', limit=10)

# Criar um DataFrame pandas com os dados
dados_zabbix = pd.DataFrame(items)

# Salvar os dados em um arquivo Excel
data_formatada = datetime.now().strftime('%d-%m-%y')
arquivo_excel = f"dados_zabbix_{data_formatada}.xlsx"

caminho_arquivo = os.path.join(fr"C:\caminho\para\o\seu\diretorio\arquivos", arquivo_excel)
dados_zabbix.to_excel(caminho_arquivo, index=False)

print(f"Arquivo '{arquivo_excel}' criado com sucesso!")

# Desconectar-se do servidor Zabbix
zapi.logout()
