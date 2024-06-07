import win32com.client as win32
from datetime import datetime
import Dados

def enviar_emil(fornecedor, dados_tabela):
    data_atual = datetime.now().strftime("%d/%m/%Y")
    outlook = win32.Dispatch('outlook.application')

    email = outlook.CreateItem(0)
    email.To = "alice.moura.101@gmail.com" #Dados.email_do_fornecedor(fornecedor)
    email.Subject = f"{data_atual} {fornecedor} - Prazos"

    linhas_tabela = ""
    for linha in dados_tabela:
        linhas_tabela += f"<tr><td>{linha[0]}</td><td>{linha[1]}</td><td>{linha[2]}</td><td>{linha[3]}</td><td>{linha[4]}</td><td>{linha[5]}</td><td>{linha[6]}</td><td>{linha[7]}</td><td>{linha[8]}</td><td>{linha[9]}</td></tr>"

    email.HTMLBody = f"""
    <p>Bom dia!</p>
 
<p>Por gentileza, referente a tabela abaixo, confirmar os prazos de entrega e chegada dos pedidos para a Bosch - Campinas:</p>
 
<p>- Solicitamos os prazos o quanto antes e com urgência por tratativas internas.</p>
<p>- Caso ocorra de não receber o PR, pedimos que sinalize para que façamos o envio novamente.</p> 
<p>- Para pedidos marcados em vermelho, por favor, informar nova data e motivo do atraso com relação ao prazo anterior.</p>

    <table border="1">
        <tr>
            <th>Data Remessa</th>
            <th>Item</th>
            <th>Pedido de Compras</th>
            <th>Material</th>
            <th>Descrição do Material</th>
            <th>Qtde. div.</th>
            <th>UM</th>
            <th>Fornecedor</th>
            <th>Razão Social</th>
            <th>Prazo</th>
        </tr>
        {linhas_tabela}
    </table>
    """

    email.Send()
    print("email enviado")
