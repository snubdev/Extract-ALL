import pandas as pd
import mysql.connector

# conexão com banco de dados
host = "host"
dbname = "dbname"
user = "user"
password = "password"
sslmode = "sslmode"

conn_string = "host={0} user={1} dbname={2} password={3} sslmode={4}".format(host, user, dbname, password, sslmode)
try:
    conn = mysql.connector.connect(conn_string)
    cursor = conn.cursor()
    print('Conectado')
except mysql.connector.Error as e:
    print(f'falha na conexão {e}')


def coletor(arquivo):

    df = pd.read_excel(arquivo)
    i = 0

    while True:
        c1 = df.loc[i, "Grupo de Exames"]
        c2 = df.loc[i, "Código"]
        c3 = df.loc[i, "Exame"]
        c4 = df.loc[i, "Data vigência"]
        c5 = df.loc[i, "Rotina"]
        c6 = df.loc[i, "Prazo de Entrega"]
        c7 = df.loc[i, "VM CARD"]
        c8 = df.loc[i, "FUNERÁRIA"]
        c9 = df.loc[i, "PARTICULAR"]

        print(f'{c1}, {c2}, {c3}, {c4}, {c5}, {c6}, {c7}, {c8}, {c9}')

        if i >= 0:
            cursor.execute(f'''INSERT INTO Nome_da_Tabela (grupo_de_exames, codigo, exame, data_vigencia,
                              rotina, prazo_de_entrega, vm_card, funeraria, particular) VALUES ('{c1}', '{c2}', '{c3}', 
                              '{c4}', '{c5}', '{c6}', '{c7}', '{c8}', '{c9}');''')

            conn.commit()
            print(f'Dados enviados com sucesso - {i}')
        else:
            print('Dados não enviados')

        i += 1

        if i == len(df):
            break

    return print('concluindo')

coletor("D:/Usuario/PycharmProjects/Extract-ALL/teste.xlsx")