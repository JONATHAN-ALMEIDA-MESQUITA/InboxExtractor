##impotando lib imap_tools para acessar conta do no email
from imap_tools import MailBox, AND, OR, A
import os


# Obtendo o valor da variável de ambiente
email = os.environ.get('EMAIL_OUT')
senha = os.environ.get('SENHA_OUT')

# logar no email utilizando as variaves de ambiente
meu_email = MailBox('imap-mail.outlook.com').login(email, senha)

#Fazendo busca no email com base no titulo e na palavra no corpo do email.
lista_email = meu_email.fetch(AND(subject='Teste 3'))

lista = ['Ola Mundo', 'Olá Mundo', 'Hello World']

# Iterar sobre os e-mails encontrados
for email in lista_email:
    # Iterar sobre cada palavra na lista
    for palavra in lista:
        # Verificar se a palavra está presente no corpo do e-mail (em minúsculas para garantir correspondência)
        if palavra.lower() in email.text.lower():
            # Se encontrarmos a palavra, imprimir algumas informações sobre o e-mail
            print("Data:", email.date)
            print("Assunto:", email.subject)
            print("Remetente:", email.from_)
            print("Destinatário:", email.to)
            print("Texto do Corpo:", email.text)
            print()  # Adicionar uma linha em branco entre os e-mails para melhorar a legibilidade
            break #Como já encontramos a palavra, podemos interromper a iteração sobre a lista