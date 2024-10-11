import pandas as pd
from datetime import datetime, timedelta


#file_path = 'C:/Users/prr8ca/Desktop/Projetos/Email/mailQT/BLOCO K - Planejamento.xlsx'
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
    df['Depósito'] = df['Depósito'].astype(str).str.strip()
    deposito_unicos = df['Depósito'].dropna().unique()
    deposito_lista = deposito_unicos.tolist()
    return deposito_lista


def email_do_fornecedor(nome_fornecedor): 
    filtro_fornecedor = df['Razão Social'] == nome_fornecedor
    email = df.loc[filtro_fornecedor, 'E-mail Fornecedor'].iloc[0] if filtro_fornecedor.any() else None
    print(email)
    return email


def ver_prazos_vencidos(nome_fornecedor, tipo):
    prazos_vencidos = []
    prazos_vazios = []
    prazos_15_dias = []
    filtro_fornecedor = df['Razão Social'] == nome_fornecedor
    for index, row in df[filtro_fornecedor].iterrows():
        excel_linha_index = index + 2  
        prazo_cell = row['Prazo']

        if isinstance(prazo_cell, pd._libs.tslibs.nattype.NaTType) or isinstance(prazo_cell, float):
            prazos_vazios.append(excel_linha_index)

        elif prazo_cell + timedelta(days=1) < datetime.now():
            prazos_vencidos.append(excel_linha_index)

        elif prazo_cell <= datetime.now() + timedelta(days=15):
            prazos_15_dias.append(excel_linha_index)
            

    if tipo == "vencidos":
        print("prazos vencidos: ", prazos_vencidos)
        return prazos_vencidos
    
    elif tipo == "vazio":
        print("prazos vazios: ", prazos_vazios)
        return prazos_vazios
    
    elif tipo == "15":
        print("prazos que faltam 15 dias: ", prazos_15_dias)
        return prazos_15_dias
        


def ver_informacoes_necessarias(numeros_linha):
    informacoes = []
    for numero_linha in numeros_linha:
        linha = df.iloc[numero_linha - 2]  

        informacao_linha = [valor.strftime("%d/%m/%Y") if isinstance(valor, pd.Timestamp) else valor for i, valor in enumerate(linha) if i in [0, 1, 2, 3, 4, 5, 6, 7, 8, 14]]
        informacoes.append(informacao_linha)
    print(informacoes)
    return informacoes

