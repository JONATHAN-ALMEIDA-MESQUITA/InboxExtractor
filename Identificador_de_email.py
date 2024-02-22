##impotando lib imap_tools para acessar conta do no email
from imap_tools import MailBox, AND, OR, A
import os

# Obtendo o valor da variável de ambiente
email = os.environ.get('EMAIL_OUT')
senha = os.environ.get('SENHA_OUT')

# logar no email utilizando as variaves de ambiente
meu_email = MailBox('imap-mail.outlook.com').login(email, senha)

lista_email = meu_email.fetch(AND(subject= 'Teste 3'))
print(len(list(lista_email)))