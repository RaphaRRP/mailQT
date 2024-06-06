import win32com.client as win32
from datetime import datetime
import Dados

def enviar_emil(fornecedor):
    data_atual = datetime.now().strftime("%d/%m/%Y")
    outlook = win32.Dispatch('outlook.application')

    email = outlook.CreateItem(0)
    email.To = "alice.moura.101@gmail.com" #Dados.email_do_fornecedor(fornecedor)
    email.Subject = f"{data_atual} {fornecedor} - Prazos"
    dados_tabela = [
        ["Valor 1, Coluna 1", "Valor 1, Coluna 2"],
        ["Valor 2, Coluna 1", "Valor 2, Coluna 2"],
        ["Valor 3, Coluna 1", "Valor 3, Coluna 2"],
    ]

    linhas_tabela = ""
    for linha in dados_tabela:
        linhas_tabela += f"<tr><td>{linha[0]}</td><td>{linha[1]}</td></tr>"

    email.HTMLBody = f"""
    <p>Código python teste</p>

    <table border="1">
        <tr>
            <th>Coluna 1</th>
            <th>Coluna 2</th>
        </tr>
        {linhas_tabela}
    </table>
    """

    email.Send()
    print("email enviado")

enviar_emil("fornecedor")