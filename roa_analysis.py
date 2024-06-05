import pandas as pd
import os

diretorio = 'C:/Users/Dev M7/Desktop/INV/rec/dra'
lista_dfs = []

colunas_para_excluir = ['Papel', 'C. Custo', 'Manual', 'C.Pagar', 'Descrição', 'Equipe Responsável', 'Unid. Negócio Resp.']
codigos_para_excluir = ['M18', 'A71632', 'A74905', 'A69713', 'M19', 'AKD1', 'M25', 'GVM', 'G99']

arquivos = os.listdir(diretorio)

for arquivo in arquivos:
    if arquivo.endswith('.xlsx'):
        caminho_arquivo = os.path.join(diretorio, arquivo)
        df = pd.read_excel(caminho_arquivo, skiprows=1)
        df = df.drop(columns=colunas_para_excluir)
        df = df[~df['Cód. Interno'].isin(codigos_para_excluir)]