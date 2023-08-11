#importações
import pandas as pd
import pathlib as pa
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

#importar e tratar base de dados
emails = pd.read_excel(r'SEU CAMINHO\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'SEU CAMINHO\Projeto AutomacaoIndicadores\Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'SEU CAMINHO\Projeto AutomacaoIndicadores\Bases de Dados\Vendas.xlsx')
#criar um arquivo pra cada loja
vendas = vendas.merge(lojas, on='ID Loja')

#Substitui todos os emails
new_email = 'email_destino@gmail.com'
emails['E-mail'] = new_email
emails.to_excel(r'SEU CAMINHO\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx', index=False)

#Dicionario para lojas
dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja,:]

#Dia da análise
dia_indicador = vendas['Data'].max()

#Cria as pastas de backup
caminho_backup = pa.Path(r'SEU CAMINHO\Projeto AutomacaoIndicadores\Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = []
for arquivo in arquivos_pasta_backup:
    lista_nomes_backup.append(arquivo.name)

for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
    nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)

#Definição de metas
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdprodutos_dia = 4
meta_qtdprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

for loja in dicionario_lojas:
    #Calcula faturamento
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    #Diversidade de produtos
    qtd_produtos_ano = len(vendas_loja['Produto'].unique())
    qtd_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    #Calcula ticket médio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    #Define a situação atual da loja
    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'green'
    if qtd_produtos_dia >= meta_qtdprodutos_dia:
        cor_qtd_dia = 'green'
    else:
        cor_qtd_dia = 'red'
    if qtd_produtos_ano >= meta_qtdprodutos_ano:
        cor_qtd_ano = 'green'
    else:
        cor_qtd_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_tic_dia = 'green'
    else:
        cor_tic_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_tic_ano = 'green'
    else:
        cor_tic_ano = 'red'

    #Função para enviar email
    def enviar_email():
        nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
        #Corpo do email
        corpo_email = f"""
    <p>Bom dia, <strong>{nome}</strong></p>
    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da loja <strong>{loja}</strong> foi:</p>

    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
    </tr>
    <tr>
        <td>Faturamento</td>
        <td style="text-align: center;">R${faturamento_dia:.2f}</td>
        <td style="text-align: center;">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center;"><font color="{cor_fat_dia}">◙</font></td>
    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center;">{qtd_produtos_dia}</td>
        <td style="text-align: center;">{meta_qtdprodutos_dia}</td>
        <td style="text-align: center;"><font color="{cor_qtd_dia}">◙</font></td>
    </tr>
    <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center;">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center;">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center;"><font color="{cor_tic_dia}">◙</font></td>
    </tr>
    
    </table>
    <br>
    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
    </tr>
    <tr>
        <td>Faturamento</td>
        <td style="text-align: center;">R${faturamento_ano:.2f}</td>
        <td style="text-align: center;">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center;"><font color="{cor_fat_ano}">◙</font></td>
    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center;">{qtd_produtos_ano}</td>
        <td style="text-align: center;">{meta_qtdprodutos_ano}</td>
        <td style="text-align: center;"><font color="{cor_qtd_ano}">◙</font></td>
    </tr>
    <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center;">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center;">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center;"><font color="{cor_tic_ano}">◙</font></td>
    </tr>
    
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>

    <p>Atenciosamente, Renan Aquino</p>
    """
        #Password do email, para autorizar apps de terceiros
        password = 'sua senha'
        
        # Cria um objeto MIMEMultipart para o email
        msg = MIMEMultipart()
        msg['Subject'] = f'OnePage dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
        msg['From'] = 'email remetente'
        msg['To'] = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
        msg.attach(MIMEText(corpo_email, 'html'))
        attachment = pa.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
        
        with open(attachment, 'rb') as file:
            attach_part = MIMEApplication(file.read(), Name=f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx')
        attach_part['Content-Disposition'] = f'attachment; filename="{attach_part.get_filename()}"'
        msg.attach(attach_part)

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        
        # Login
        s.login(msg['From'], password)
        s.sendmail(msg['From'], msg['To'].split(','), msg.as_string().encode('utf-8'))
        print(f'Email loja {loja} enviado!')

    #Enviar email
    enviar_email()

#Ranking de lojas
faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
faturamento_lojas_ano.to_excel(r'{} / {}'.format(caminho_backup, nome_arquivo))

vendas_dia = vendas.loc['Data'==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
faturamento_lojas_dia.to_excel(r'{} / {}'.format(caminho_backup, nome_arquivo))

#Enviar email para diretoria
def email_diretoria():
    #Corpo do email
    corpo_email = f"""
    Prezados, bom dia!

    Melhor loja do dia em faturamento: Loja {faturamento_lojas_dia.index[0]} com faturamento R${faturamento_lojas_dia.iloc[0, 0]}:.2f
    Pior loja do dia em faturamento: Loja {faturamento_lojas_dia.index[-1]} com faturamento R${faturamento_lojas_dia.iloc[-1, 0]}:.2f

    Melhor loja do ano em faturamento: Loja {faturamento_lojas_ano.index[0]} com faturamento R${faturamento_lojas_ano.iloc[0, 0]}:.2f
    Pior loja do ano em faturamento: Loja {faturamento_lojas_ano.index[-1]} com faturamento R${faturamento_lojas_ano.iloc[-1, 0]}:.2f

    Segue em anexo o ranking do ano e do dia de todas as lojas.

    Qualquer dúvida estou à disposição.

    Atenciosamente,
    Renan Aquino
    """
    #Password do email, para autorizar apps de terceiros
    password = 'sua senha'

    # Cria um objeto MIMEMultipart para o email

    msg = MIMEMultipart()
    msg['Subject'] = f'Ranking dia {dia_indicador.day}/{dia_indicador.month}'
    msg['From'] = 'email remetente'
    msg['To'] = emails.loc[emails['Loja'] == 'Diretoria', 'E-mail'].values[0]
    msg.attach(MIMEText(corpo_email, 'html'))
    attachment = pa.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
    with open(attachment, 'rb') as file:
        attach_part = MIMEApplication(file.read(), Name=f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx')
    attach_part['Content-Disposition'] = f'attachment; filename="{attach_part.get_filename()}"'
    msg.attach(attach_part)

    attachment = pa.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
    with open(attachment, 'rb') as file:
        attach_part = MIMEApplication(file.read(), Name=f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx')
    attach_part['Content-Disposition'] = f'attachment; filename="{attach_part.get_filename()}"'
    msg.attach(attach_part)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()

    # Login
    s.login(msg['From'], password)
    s.sendmail(msg['From'], msg['To'].split(','), msg.as_string().encode('utf-8'))
    print(f'Email Diretoria enviado!')
email_diretoria()