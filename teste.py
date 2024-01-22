
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

def obter_conta_por_endereco(outlook, endereco):
    accounts = outlook.Session.Accounts
    for account in accounts:
        if account.SmtpAddress.lower() == endereco.lower():
            return account
        print(account)
    return None

remetente = 'contasapagar@webmotors.com.br'
conta = obter_conta_por_endereco(outlook, remetente)

if conta:
    print(f'Remetente encontrado: {conta.SmtpAddress}')
else:
    print('Remetente não encontrado.')
