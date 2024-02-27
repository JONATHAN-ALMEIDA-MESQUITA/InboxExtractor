import os
import re
import pandas as pd
from imap_tools import MailBox, AND
from fpdf import FPDF

base = pd.read_excel('base test.xlsx')
print(base)
print('-' * 60)

# Obtendo o valor da variável de ambiente
email = os.environ.get('EMAIL_OUT')
senha = os.environ.get('SENHA_OUT')

# Logar no email utilizando as variáveis de ambiente
meu_email = MailBox('imap-mail.outlook.com').login(email, senha)

# Criar uma lista para armazenar os resultados
resultados = []

# Expressão regular para buscar 'Não (x)' ignorando espaços, letras maiúsculas/minúsculas e acentos
regex_nao = re.compile(r'n[ãaáàâäeéèêëiíìîïoóòôöuúùûü]*\s*\(\s*x\s*\)', re.IGNORECASE)

# Criar uma pasta para salvar os e-mails comissionados
pasta_salvar = 'emails_comissionados'
if not os.path.exists(pasta_salvar):
    os.makedirs(pasta_salvar)

# Iterar sobre os vouchers da base de dados
for voucher in base['Voucher']:
    # Buscar e-mails com base no voucher
    lista_email = list(meu_email.fetch(AND(subject=str(voucher))))

    # Exibir informações para validação
    print('-' * 60)
    print(f"Voucher: {voucher}")
    print(f"Tamanho da lista de e-mails para o voucher {voucher}: {len(lista_email)}")

    # Verificar se houve algum e-mail encontrado
    if lista_email:
        # Exibir informações de cada e-mail
        for email in lista_email:
            print("Data:", email.date)
            print("Assunto:", email.subject)
            print("Remetente:", email.from_)
            print("Destinatário:", email.to)
            print("Texto do Corpo:", email.text)
            print()  # Adicionar uma linha em branco entre os e-mails para melhorar a legibilidade

        # Verificar se a palavra-chave 'Sim (x)' está presente nos e-mails
        encontrado_sim = any('Sim (x)' in email.text for email in lista_email)
        # Verificar se a expressão regular é encontrada em algum e-mail
        encontrado_nao = any(regex_nao.search(email.text) for email in lista_email)

        # Determinar o status com base nas palavras-chave encontradas nos e-mails
        if encontrado_sim:
            resultados.append('Comissionado')
            # Salvar o e-mail como arquivo PDF se estiver comissionado
            for idx, email in enumerate(lista_email):
                # Nome do arquivo para salvar o e-mail com o número do voucher
                nome_arquivo = f'email_{voucher}_{idx + 1}.pdf'
                caminho_arquivo = os.path.join(pasta_salvar, nome_arquivo)
                # Criar o PDF
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                # Adicionar o conteúdo do e-mail ao PDF
                pdf.cell(200, 10, txt=email.subject, ln=True, align='C')
                pdf.cell(200, 10, txt=f"Remetente: {email.from_}", ln=True)
                pdf.cell(200, 10, txt=f"Destinatário: {email.to}", ln=True)
                pdf.cell(200, 10, txt="Texto do Corpo:", ln=True)
                pdf.multi_cell(0, 10, txt=email.text)
                # Salvar o PDF
                pdf.output(caminho_arquivo)
        elif encontrado_nao:
            resultados.append('Não Comissionado')
        else:
            resultados.append('Encontrado, sem confirmação')
    else:
        # Se o voucher não foi encontrado na lista de e-mails
        resultados.append('Não encontrado')

# Adicionar os resultados como uma nova coluna ao DataFrame
base['Status'] = resultados

# Exibir o DataFrame atualizado
print('-' * 60)
print(base)