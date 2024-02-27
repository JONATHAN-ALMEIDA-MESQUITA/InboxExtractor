import os
import re
import pandas as pd
from imap_tools import MailBox, AND, OR

base = pd.read_excel('C:/Projeto leitor de e-mail/Arquivos/base test.xlsx')
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