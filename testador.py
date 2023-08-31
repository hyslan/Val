# testador.py
'''Área de testes.'''
# from sqlalchemy import URL, create_engine, select, MetaData, Table, and_

# url_object = URL.create(
#     drivername="mssql+pyodbc",
#     host="10.66.9.46",
#     database="BD_MLG",
#     query={"driver": "ODBC Driver 18 for SQL Server",
#            "TrustServerCertificate": "yes",
#            "authentication": "ActiveDirectoryIntegrated", },
# )
# engine = create_engine(url_object)
# conn = engine.connect()
# metadata = MetaData(schema='LESTE_AD\\hcruz_novasp')
# # table = Table(
# #     'tbHyslancruz_Parametros',
# #     metadata,
# #     autoload_replace=True,
# #     autoload_with=engine
# # )
# # stmt = select([table])

# # results = conn.execute(stmt)
# # print(results)

import winsound

duration = 100
freq = 100
# winsound.Beep(freq, duration)
i = 0
for i in range(i, 1000):
    winsound.Beep(freq, duration)
    freq += 10
    i += 1
