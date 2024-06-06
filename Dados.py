import pandas as pd

file_path = 'C:/Users/prr8ca/Desktop/mFornecedor/BLOCO K - Planejamento.xlsx'
df = pd.read_excel(file_path, sheet_name='FUP Prazos', engine='openpyxl')

def buscar_fornecedores(): #Retorna uma lista com todos os fornecedores

    razao_social_unicos = df['Razão Social'].unique()
    razao_social_lista = razao_social_unicos.tolist()
    return razao_social_lista

def email_do_fornecedor(nome_fornecedor): #Retorna uma Lista com os email do fornecedor
    todos_emails = []
    filtro_fornecedor = df['Razão Social'] == nome_fornecedor
    email = df.loc[filtro_fornecedor, df.columns[26]].iloc[0] if filtro_fornecedor.any() else None

    if ';' in str(email):
        emails_separados = [e.strip() for e in email.split(';')]
        todos_emails.extend(emails_separados)
    elif email:
        todos_emails.append(email.strip())
    
    return email

print(email_do_fornecedor("Bosch Rexroth Ltda."))