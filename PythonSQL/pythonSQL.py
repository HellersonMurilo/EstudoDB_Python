import pyodbc

# dados de conexao
dados_conexao = (
    "Driver={SQL SERVER};"
    "Server=DSN1381250990;"
    "Database=Códigos;"
    "Trusted_Connection=yes;"  # Exemplo de conexão confiável, você pode precisar ajustar de acordo com suas configurações
)

# CONEXÃO COM O BANCO DE DADOS
try:
    conexao = pyodbc.connect(dados_conexao)
    print(f"Conexão bem-sucedida")
except pyodbc.Error as e:
    print(f"Houve um erro na sua conexão com o banco de dados: {e}")

try:
    # cursor é o método que permite realizar Querys
    cursor = conexao.cursor()

    items = []
    for i in range(1,6):
        temp = input("Informe um código")
        items.append(temp)

    comando_sql = f'''INSERT INTO Numeros(valor1, valor2, valor3, valor4, valor5)
    VALUES ({", ".join(map(str, items))})'''

    print(comando_sql)

    cursor.execute(comando_sql)
    #commit é utilizado quando precisamos realizar uma atualização no DB
    cursor.commit()
except pyodbc.Error as error_comand:
    print(f'Erro ao executar a query: {error_comand}')