import pandas as pd
from datetime import datetime, timedelta


file_path = 'S:/PM/ter/tef/tef3/Inter_Setor/TEF3 Controles/Bloco K Planejamento/BLOCO K - Planejamento.xlsx'


df = pd.read_excel(file_path, sheet_name='FUP Prazos', engine='openpyxl')


def buscar_fornecedores(depositos=None):
    if depositos is not None:
        depositos = [str(deposito) for deposito in depositos]
        fornecedores_filtrados = df[df['Depósito'].astype(str).isin(depositos)]
    else:
        fornecedores_filtrados = df

    razao_social_unicos = fornecedores_filtrados['Razão Social'].unique()
    razao_social_lista = razao_social_unicos.tolist()
    return razao_social_lista


def buscar_depositos():
    deposito_unicos = df['Depósito'].unique()
    deposito_lista = deposito_unicos.tolist()
    deposito_strings = [dep for dep in deposito_lista if isinstance(dep, str)]
    return deposito_strings


def email_do_fornecedor(nome_fornecedor): 
    filtro_fornecedor = df['Razão Social'] == nome_fornecedor
    email = df.loc[filtro_fornecedor, 'E-mail Fornecedor'].iloc[0] if filtro_fornecedor.any() else None
    return email


def ver_prazos_vencidos(nome_fornecedor):
    prazos_vencidos = []
    filtro_fornecedor = df['Razão Social'] == nome_fornecedor
    for index, row in df[filtro_fornecedor].iterrows():
        excel_linha_index = index + 2  
        prazo_cell = row['Prazo']
        

        if prazo_cell is None or prazo_cell == '' or not isinstance(prazo_cell, pd.Timestamp) or prazo_cell + timedelta(days=1) < datetime.now():
            prazos_vencidos.append(excel_linha_index)
        
    return prazos_vencidos


def ver_prazos_15_dias_para_vencer(nome_fornecedor):
    prazos_para_vencer = []
    filtro_fornecedor = df['Razão Social'] == nome_fornecedor
    for index, row in df[filtro_fornecedor].iterrows():
        excel_linha_index = index + 2  
        prazo_cell = row['Prazo']  
        cell_content = prazo_cell
        if cell_content is not None and datetime.now() <= cell_content <= datetime.now() + timedelta(days=15):
            prazos_para_vencer.append(excel_linha_index)
    return prazos_para_vencer


def ver_informacoes_necessarias(numeros_linha):
    informacoes = []
    for numero_linha in numeros_linha:
        linha = df.iloc[numero_linha - 2]  

        informacao_linha = [valor.strftime("%d/%m/%Y") if isinstance(valor, pd.Timestamp) else valor for i, valor in enumerate(linha) if i in [0, 1, 2, 3, 4, 5, 6, 7, 8, 13]] 
        informacoes.append(informacao_linha)
    return informacoes

