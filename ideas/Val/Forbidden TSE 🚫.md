# TSEs leave out the plumbing services by Iara.

        chief_iara_orders = "'534200', '534300', '537000', '537100', '538000'" if contrato == "4600042975" else "'',"

        sql_command = ("SELECT Ordem, COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] "

                       f"WHERE FAMILIA IN ({family_str}) AND Contrato = '{

                           contrato}' "

                       f"AND TSE_OPERACAO_ZSCP NOT IN ( "

                       "'731000', '733000', '743000', '745000', '785000', '785500', "  # -- SERVIÇOS DE ASFALTO

                       "'755000', '714000', '782500', '282000', '300000', '308000', '310000', '311000', '313000', "

                       "'315000', '532000', '564000', '588000', '590000', '709000', '700000', '593000', '253000', "

                       "'250000', '209000', '605000', '605000', '263000', '255000', '254000', '282000', '265000', "

                       "'260000', '265000', '263000', '262000', '284500', '286000', '282500', "  # -- RAMAL ÁGUA

                       # UNITÁRIO

                       # Obeying Iara's orders

                       f"{chief_iara_orders}"

                       "'136000', '159000', '155000') "  # -- CRIAR LÓGICA

                       )