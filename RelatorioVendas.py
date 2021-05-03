"""
Enviando Relatório de Vendas de cada loja por email (usando outloook)
Criação: ETLS em 03/05/2021

Esse código é uma adaptação minha do aprendizado do Minicurso Python da hashtag programação.
link do curso: https://pages.hashtagtreinamentos.com/minicurso-python-automacao?blog=1n4033rer&video=3dep762tr

Requisitos: 
• instalar as biliotecas do python (pandas, win32com e pyautogui)
    . para isso use o cmd e rode esses comandos
        pip install pandas
        pip install win32com
        pip install pyautogui
• Ter o outlook instalado e configurado no email do remetente.
"""

# Importanto as bibliotecas 
import pandas as pd
import win32com.client as win32
import pyautogui

class RelatorioVendas:
    
    def __init__(self):
        """ variáveis iniciais. """
        
        self.tabela_vendas = ''
        self.faturamento = ''
        self.qtq_produtos = ''
        self.ticket_medio = ''
        self.destinatario = 'dudahmovies@gmail.com' # email para quem quer enviar
    
    def base_dados(self):
        """ importar a base de dados. """
        
        self.tabela_vendas = pd.read_excel('Vendas.xlsx') # Vendas.xlsx -> base de dados utilizada nesse exemplo
        
    def visualidar_base_dados(self):
        """ visualizar a base de dados. """
        
        pd.set_option('display.max_columns', None) # para ler todas as colunas 
        print(self.tabela_vendas)
        print('-' * 50)
        
    def filtrar_tab_faturamento(self):
        """ faturamento por loja. """
        
        self.faturamento = self.tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum() # filtrando e agrupando coluna
        # print(self.faturamento)
        # print('-' * 50)
        
    def filtrar_tab_quantidade_produtos(self):
        """ quantidade de produtos vendidos por loja. """
        
        self.qtq_produtos = self.tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum() # filtrando e agrupando coluna
        # print(self.qtq_produtos)
        # print('-' * 50)
        
    def tab_ticket_medio(self):
        """ ticket médio por produto em cada loja  (faturamento / quantidade). """
        
        self.ticket_medio = (self.faturamento['Valor Final'] / self.qtq_produtos['Quantidade']).to_frame() # dividindo valores de colunas e transformando o resultado da divisão em outra coluna
        self.ticket_medio = self.ticket_medio.rename(columns={0: 'Ticket Médio'})
        # print(self.ticket_medio)
        # print('-' * 50)
        
    def enviar_email(self):
        """ enviar email com relatório. """
        
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = self.destinatario
        mail.Subject = 'Relatório de Vendas por Loja'
        mail.HTMLBody = '''
        <p>Prezados,</p>

        <p>Segue o Relatório de Vendas por cada Loja.</p>

        <p>Faturamento:</p>
        {}

        <p>Quantidade Vendida:</p>
        {}

        <p>Ticket Médio dos Produtos em cada Loja:</p>
        {}

        <p>Qualquer dúvida estou à disposição.</p>

        <p>Att.,</p>
        <p>Eduarda</p>
        '''.format(self.faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format}), self.qtq_produtos.to_html(), self.ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format}))

        mail.Send()

        pyautogui.alert('Email enviado com sucesso!')


try:
    scrip = RelatorioVendas()
    scrip.base_dados()
    # scrip.visualidar_base_dados()
    scrip.filtrar_tab_faturamento()
    scrip.filtrar_tab_quantidade_produtos()
    scrip.tab_ticket_medio()
    scrip.enviar_email()
except:
    pyautogui.alert('Erro na execução')
finally:
    import os
    os.system('pause')