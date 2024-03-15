# Importe as bibliotecas necessárias
import PySimpleGUI as sg
import pyodbc
import pandas as pd
from dotenv import load_dotenv
import os

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

class InterfaceGrafica:
    def __init__(self):
        self.items = []

    def iniciar(self):
        layout = [
            [sg.Text('Informe os códigos e clique em "Adicionar":')],
            [sg.InputText(key='-INPUT-')],
            [sg.Listbox(values=[], size=(30, 5), key='-LISTBOX-')],
            [sg.Button('Adicionar')],
            [sg.Button('Submit')]
        ]

        self.window = sg.Window('Lista Separada por Espaços', layout)

        while True:
            event, values = self.window.read()

            if event == sg.WINDOW_CLOSED:
                break
            elif event == 'Adicionar':
                input_text = values['-INPUT-']
                if input_text.strip() != '':
                    self.items.extend(input_text.split())
                    self.window['-LISTBOX-'].update(values=self.items)
            elif event == 'Submit':
                sg.popup('Itens selecionados:', self.items)
                # Exemplo de consulta
                query = f"SELECT * FROM [PLK_B2B].[T_RPA_CONDICAOCOMERCIAL] WHERE column_name IN ({', '.join(map(str, self.items))})"
                print(query)
                
                #Chamando o banco de dados e seus métodos
                banco_de_dados = BancoDeDados()
                banco_de_dados.executar_query(query)
                break

        self.window.close()

class BancoDeDados:
    # DADOS DE CONEXAO
    def __init__(self):
        self.server = os.getenv('SERVER')
        self.database = os.getenv('DATABASE')
        self.username = os.getenv('USERNAME_DB')
        self.password = os.getenv('PASSWORD')
        self.connection_success = False

    def conectar(self):
        # String de conexão
        connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'

        try:
            # Conecte-se ao banco de dados
            self.conn = pyodbc.connect(connection_string)
            self.connection_success = True
        except pyodbc.Error as e:
            sg.Popup(f'Erro de conexão com o banco de dados: {e}')

    def executar_query(self, query):
        if self.connection_success:
            try:
                # Lê os dados do banco de dados usando pandas
                df = pd.read_sql(query, self.conn)

                # Salva os dados em um arquivo Excel
                nome_arquivo_excel = 'dados_do_banco.xlsx'
                df.to_excel(nome_arquivo_excel, index=False)

                sg.Popup(f'Os dados foram salvos no arquivo Excel: {nome_arquivo_excel}')

            except pyodbc.Error as e:
                sg.Popup(f'Erro ao executar a query: {e}')
        else:
            sg.Popup('Não foi possível executar a query, pois a conexão com o banco de dados não foi estabelecida.')

if __name__ == "__main__":
    banco_de_dados = BancoDeDados()
    banco_de_dados.conectar()
    if banco_de_dados.connection_success:
        interface_grafica = InterfaceGrafica()
        interface_grafica.iniciar()
