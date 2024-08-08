import win32com.client as win32
from datetime import datetime
import pandas as pd
import Dados


def enviar_email(fornecedor, dados_tabela_vencido, dados_tabela_15, dados_tabela_vazio):
    data_atual = datetime.now().strftime("%d/%m/%Y")
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    email.To = f"{Dados.email_do_fornecedor(fornecedor)};lucas.martinez@br.bosch.com"
    #email.To = "alice.moura.101@gmail.com"


    
    email.Subject = f"[{data_atual}] {fornecedor} - Prazos"


    def formatar_linhas_tabela(dados_tabela, tipo):
        linhas_tabela = ""
        for linha in dados_tabela:
            data_str = str(linha[9])

            try:
                data_str = datetime.strptime(data_str, "%Y-%m-%d %H:%M:%S")
                data_str = data_str.strftime("%d/%m/%Y")
            except Exception as e:
                ...

            if tipo == "vencidos":
                estilo_celula = 'background-color: rgb(255, 0, 0);'

            elif tipo == "15":
                estilo_celula = 'background-color: rgb(0, 255, 0);'

            elif tipo == "vazio":
                estilo_celula = 'background-color: rgb(255, 255, 0); color: rgb(255, 255, 0);'

            linhas_tabela += f"<tr>"
            for item in linha[:9]:
                linhas_tabela += f"<td>{item}</td>"
            linhas_tabela += f"<td style='{estilo_celula}'>{data_str}</td></tr>"
        return linhas_tabela


    linhas_tabela_vencido = formatar_linhas_tabela(dados_tabela_vencido, "vencidos")
    linhas_tabela_vazio = formatar_linhas_tabela(dados_tabela_vazio, "vazio")
    linhas_tabela_15 = formatar_linhas_tabela(dados_tabela_15, "15")
    
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
        {linhas_tabela_vazio}
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
