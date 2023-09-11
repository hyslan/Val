from urllib.parse import quote_plus
from sqlalchemy import create_engine, text
import pandas as pd

server_name = '10.66.42.188'
database_name = 'BD_MLG'
connection_string = f'DRIVER={{SQL Server Native Client 11.0}};SERVER={server_name};DATABASE={database_name};Trusted_Connection=yes;'
encoded_connection_string = quote_plus(connection_string)
connection_url = f"mssql+pyodbc:///?odbc_connect={encoded_connection_string}"
engine = create_engine(connection_url)
tse = '215000'

sql_command2 = (
    "SELECT Ordem FROM [LESTE_AD\hcruz_novasp].[v_Hyslan_Engetami_Valoracao] WHERE TSE_OPERACAO_ZSCP = '{}'").format(tse)
# sql_command2 = ("UPDATE [LESTE_AD\maguiar_tejofra].[TB_DADOS_REFAT_FAZER] SET {} = '{}' WHERE {} = '{}'").format(
#     str(nome_coluna), str(valor_a_ser_atualizado), str(coluna_com_cod_unico), str(valor_cod_unico))
with engine.connect() as cnn:
    teste2 = cnn.execute(text(sql_command2))
    cnn.commit()
    df = pd.read_sql(sql_command2, engine)
    cnn.close()
df = df.reset_index()
df_list = df['Ordem'].tolist()
for ordem in df_list:
    print(ordem)
