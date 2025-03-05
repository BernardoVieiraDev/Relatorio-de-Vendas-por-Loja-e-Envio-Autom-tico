# Importação das bibliotecas necessárias
import pandas as pd  # Biblioteca para manipulação de dados em formato de tabelas (DataFrames)
import win32com.client as win32  # Biblioteca para automação de tarefas no Windows, incluindo o envio de e-mails via Outlook

# Definição da classe CalcularFaturamento
class CalcularFaturamento:
    
    # Construtor da classe, onde o arquivo Excel é carregado para análise
    def __init__(self, tabela_vendas=None):
        # Caso o parâmetro 'tabela_vendas' não seja fornecido, o arquivo 'Vendas.xlsx' é carregado
        if tabela_vendas is None:
            tabela_vendas = pd.read_excel('Vendas.xlsx')  # Carrega o arquivo Excel com os dados de vendas
        self.tabela_vendas = tabela_vendas  # Armazena o DataFrame de vendas na instância da classe
        pd.set_option('display.max_columns', None)  # Configura a exibição para mostrar todas as colunas no DataFrame

    # Método para calcular a quantidade total de produtos vendidos por loja
    def obter_quantidade_por_loja(self):
        # Agrupa os dados pela 'ID Loja' e soma a coluna 'Quantidade' para cada loja
        self.quantidade_vendidos = self.tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
        # Extrai apenas a coluna 'Quantidade' após o agrupamento
        self.quantidade_vendidos = self.quantidade_vendidos['Quantidade']
        return self.quantidade_vendidos  # Retorna a quantidade de produtos vendidos por loja

    # Método para calcular o faturamento total por loja
    def obter_faturamento_por_loja(self):
        # Agrupa os dados pela 'ID Loja' e soma a coluna 'Valor Final' para cada loja
        self.faturamento = self.tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
        # Extrai apenas a coluna 'Valor Final' após o agrupamento
        self.faturamento = self.faturamento['Valor Final']
        return self.faturamento  # Retorna o faturamento total por loja

    # Método para calcular o faturamento total por loja (com outra implementação mais genérica)
    def calcular_faturamento_total_por_loja(self):
        # Agrupa os dados pela 'ID Loja' e soma a coluna 'Valor Final' para cada loja
        self.faturamento_loja = self.tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
        return self.faturamento_loja  # Retorna o faturamento total por loja

    # Método para calcular a quantidade total de produtos vendidos por loja (mais genérico)
    def calcular_quantidade_vendida_total_por_loja(self):
        # Agrupa os dados pela 'ID Loja' e soma a coluna 'Quantidade' para cada loja
        self.produtos_vendidos = self.tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
        return self.produtos_vendidos  # Retorna a quantidade de produtos vendidos por loja

    # Método para calcular o ticket médio por loja (faturamento dividido pela quantidade)
    def calcular_ticket_medio_por_loja(self):
        # Obtém o faturamento por loja
        self.vendas = self.obter_faturamento_por_loja()
        # Obtém a quantidade de produtos vendidos por loja
        self.quantidade = self.obter_quantidade_por_loja()
        # Calcula o ticket médio dividindo o faturamento pela quantidade
        self.ticket_medio = (self.vendas / self.quantidade).to_frame()  # Cria um DataFrame para exibir o ticket médio
        return self.ticket_medio  # Retorna o ticket médio por loja

    # Método para enviar o relatório por e-mail utilizando Outlook
    def enviar_email(self, destinatario):
        try:
            # Obtendo os dados necessários para o corpo do e-mail
            faturamento = self.obter_faturamento_por_loja()  # Chama o método para obter o faturamento
            quantidade = self.obter_quantidade_por_loja()  # Chama o método para obter a quantidade
            
            # Converte os dados em formato de string para fácil inclusão no corpo do e-mail
            faturamento_str = faturamento.to_string()
            quantidade_str = quantidade.to_string()
            
            # Criando o e-mail no Outlook
            self.outlook = win32.Dispatch("Outlook.Application")  # Inicializa o Outlook
            self.mail = self.outlook.CreateItem(0)  # Cria um novo item de e-mail
            self.mail.To = destinatario  # Define o destinatário do e-mail
            self.mail.Subject = 'Relatório de Vendas por Loja'  # Define o assunto do e-mail
            self.mail.HTMLBody = f'''
            <p>Prezados,</p>

            <p>Segue o Relatório de Vendas por Loja:</p>

            <p><strong>Faturamento:</strong><br>{faturamento_str}</p>
            
            <p><strong>Quantidade Vendidas:</strong><br>{quantidade_str}</p>
            '''
            # Envia o e-mail com os dados formatados no corpo
            self.mail.Send()
            print("E-mail enviado com sucesso!")  # Mensagem de confirmação no console

        except Exception as e:
            # Caso ocorra algum erro, exibe a mensagem de erro no console
            print(f"Ocorreu um erro ao enviar o e-mail: {e}")

# Criando a instância da classe e enviando o e-mail
calcular_faturamento = CalcularFaturamento()  # Instancia a classe 'CalcularFaturamento', carregando os dados de vendas # Envia o relatório de vendas por e-mail para o destinatário fornecido
print([m for m in dir(CalcularFaturamento) if not m.startswith('__')])
