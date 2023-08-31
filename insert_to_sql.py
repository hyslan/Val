'''Módulo de Insert para SQL'''
import pandas as pd
from sqlalchemy.engine import URL

connection_string = 'Driver={SQL Server Native Client 11.0};Server=10.66.9.46;Integrated Security=SSPI;Database=BD_MLG;'
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url)
nome_tabela = 'TB_NETA_PROV' 
schema = 'LESTE_AD\maguiar_tejofra'
dados_205 = pd.DataFrame(# colocar algum dado) 
dados_205.to_sql(con=engine, name=nome_tabela,schema=schema, if_exists='replace',index=False)