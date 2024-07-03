import win32com.client as win32
from datetime import datetime
import pandas as pd
import Dados


def colorir_celula_prazo(data_prazo):
    if pd.isnull(data_prazo) or data_prazo == '':
        return 'background-color: rgb(255, 255, 0); color: rgb(255, 255, 0);'
    
    try:
        data_prazo_datetime = datetime.strptime(data_prazo, "%d/%m/%Y")
    except ValueError:
        return 'background-color: rgb(255, 255, 0); color: rgb(255, 255, 0);'
    
    if data_prazo_datetime < datetime.now():
        return 'background-color: rgb(255, 0, 0);'
    else:
        return 'background-color: rgb(0, 255, 0);'


def enviar_email(fornecedor, dados_tabela_vencido, dados_tabela_15):
    data_atual = datetime.now().strftime("%d/%m/%Y")
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    email.To = f"{Dados.email_do_fornecedor(fornecedor)};lucas.martinez@br.bosch.com"
    #email.To = "alice.moura.101@gmail.com"
    email.Subject = f"[{data_atual}] {fornecedor} - Prazos"


    def formatar_linhas_tabela(dados_tabela):
        linhas_tabela = ""
        for linha in dados_tabela:
            data_str = str(linha[9])
            estilo_celula = colorir_celula_prazo(data_str)
            linhas_tabela += f"<tr>"
            for item in linha[:9]:
                linhas_tabela += f"<td>{item}</td>"
            linhas_tabela += f"<td style='{estilo_celula}'>{data_str}</td></tr>"
        return linhas_tabela


    linhas_tabela_vencido = formatar_linhas_tabela(dados_tabela_vencido)
    linhas_tabela_15 = formatar_linhas_tabela(dados_tabela_15)
    email.HTMLBody = f"""
    <p style="font-family: Arial, sans-serif; font-size: 10pt;">Bom dia!</p>
 
    <p style="font-family: Arial, sans-serif; font-size: 10pt;">Por gentileza, referente a tabela abaixo, confirmar os prazos de <u>entrega e chegada dos pedidos para a Bosch - Campinas</u>:</p>
 
    <p style="margin: 0; font-family: Arial, sans-serif; font-size: 9pt; font-style: italic;">- Solicitamos os prazos o quanto antes e com urgência por tratativas internas.</p>
    <p style="margin: 0; font-family: Arial, sans-serif; font-size: 9pt; font-style: italic;">- Caso ocorra de não receber o PR, pedimos que sinalize para que façamos o envio novamente.</p> 
    <p style="margin: 0; font-family: Arial, sans-serif; font-size: 9pt; font-style: italic;">- Para pedidos com prazo vencido ou sem prazo, por favor, informar nova data e motivo do atraso com relação ao prazo anterior.</p>

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
        {linhas_tabela_vencido}
    </table>

    <p style="font-family: Arial, sans-serif; font-size: 10pt;">- Para pedidos dentro de 15 dias, confirmar os prazos.</p>

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
        {linhas_tabela_15}
    </table>

    <p style="font-family: Arial, sans-serif; font-size: 10pt;">Obrigado!</p> 
    """

    email.Send()
