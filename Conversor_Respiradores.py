import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import csv
import numpy as np
import locale
from datetime import date


class Conversor:
    #variáveis mínimas para fazer uma conversão
    def __init__(self, excel_header=[],items_to_remove=[np.nan], *args, **kargs):
        root = tk.Tk()
        root.withdraw()
        self.file_path = filedialog.askopenfilename()
        excel_filename = os.path.split(self.file_path)[1]
        self.csv_filename = excel_filename.split('xls')[0] + 'csv'
        self.excel_path = self.file_path
        self.csv_filename = self.csv_filename
        self.excel_filename = excel_filename
        self.excel_file = pd.ExcelFile(self.file_path)
        self.excel_header = excel_header
        self.items_to_remove = items_to_remove
        
    @property
    def to_csv(self, loc='', *args, **kargs):
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        df =  pd.concat(pd.read_excel( (self.excel_path) , sheet_name=None),ignore_index= True)
        df =  df.replace(self.items_to_remove,np.nan)
        df =  df.dropna()
        df.to_csv(self.csv_filename, encoding='utf-8',index=False,sep=';', line_terminator='\n')


class Tabela_respiradores(Conversor):
    def __init__(self):
        #Cria um objeto do tipo conversor podendo sobreescrever os valores pré-estabelecidos pela classe
        Conversor.__init__(self, ['DATA', 'FORNECEDOR', 'DESTINO', 'ESTADO/MUNICIPIO', 'TIPO', 'QUANTIDADE', 'VALOR', 'DESTINATARIO', 'UF', 'DATA DE ENTREGA'])
             
    @property
    def to_csv(self):
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        #sheet_name=none .... indica o uso de todas as sheets
        #names .............. indica o formato do header manualmente
        #usecols............. indica as colunas que serão lidas
        #skiprows[0]......... pula o titulo de cada sheet
        df = pd.concat(pd.read_excel((self.file_path), sheet_name=None, names=self.excel_header,usecols=[1,2,3,4,5,6,7,8,9,10], 
                                    skiprows=[0]), ignore_index=True)
        df = df.replace(self.items_to_remove, np.nan)
        nrows = len(df.index) - 2
        #loop que formata as datas e seta nulo valores posteriores ao dia de pesquisa
        for i in range(nrows):
            try:
                if date.fromisoformat(str(df['DATA DE ENTREGA'][i]).split(' ')[0]) > date.today():
                    #seta nulo uma linha em que a data seja posterior ao dia de pesquisa
                    for col in self.excel_header:
                        df[col][i] = np.nan
                    continue
                #formatação de datas validas
                df['DATA DE ENTREGA'][i] = df['DATA DE ENTREGA'][i].date().strftime(
                    "%d/%m/%Y")
                df['DATA'][i] = str(df['DATA'][i].date().strftime("%d/%m/%Y"))
                df['VALOR'][i] = locale.format_string("%.2f", df['VALOR'][i], 0)
            except:
                pass
        #exclui todos os valores inválidos
        df = df.dropna()
        df.to_csv(self.csv_filename, encoding='utf-8', index=False, sep=';', line_terminator='\n')

class Tabela_EPI(Conversor):
    def __init__(self):
        Conversor.__init__(self, excel_header=['Material', 'Dt.Saída', 'Nº Pedido', 'Requisitante / Destino', 'Unidade', 'Quantidade', 'Status'])
    
    @property
    def to_csv(self):
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        df = pd.concat(pd.read_excel((self.excel_path), sheet_name=['Dados Brutos'], skiprows=[0], usecols=self.excel_header), ignore_index=True)
        df = df.replace(self.items_to_remove, np.nan)
        df = df.dropna()
        nrows = len(df.index) - 2
        for i in range(nrows):
            try:
                df['Quantidade'][i] = locale.format_string(
                    "%.2f", df['Quantidade'][i], 0)
            except:
                pass
        df = df.dropna()
        return df.to_csv(self.csv_filename, encoding='utf-8', index=False, sep=';', line_terminator='\n')


if __name__ ==  '__main__':
    conversor = Tabela_respiradores()
    conversor.to_csv