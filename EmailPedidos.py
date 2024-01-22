from Imports import *


def EmailPedidos():
    
    def enviar_email(destinatario, assunto, corpo, anexo):
        outlook = win32.Dispatch('outlook.application')

        oacctouse = None
        for oacc in outlook.Session.Accounts:
             if oacc.SmtpAddress == "contasapagar@webmotors.com.br":
                oacctouse = oacc
                break
             
        mail = outlook.CreateItem(0)
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, oacctouse))
        mail.To = destinatario
        mail.Subject = assunto
        mail.Body = corpo
        mail.Send()
        mail.CC = "contasapagar@webmotors.com.br"
        

    diretorio_pedidos = os.path.join(os.getcwd(), f'Pedidos_Enviados_{datetime.now().strftime("%d%m%Y")}')

    if os.path.exists(diretorio_pedidos):
        arquivos_excel = [arquivo for arquivo in os.listdir(diretorio_pedidos) if arquivo.endswith(".xlsx")]

        for arquivo_excel in arquivos_excel:
            nome_pessoa = os.path.splitext(arquivo_excel)[0]
            caminho_arquivo = os.path.join(diretorio_pedidos, arquivo_excel)

            if os.path.exists(caminho_arquivo):
                df_emaildb = pd.read_excel(os.path.join(os.getcwd(), "EmailDB.xlsx"))

                if nome_pessoa in df_emaildb['Name'].values:
                    email_destinatario = df_emaildb.loc[df_emaildb['Name'] == nome_pessoa, 'Contact Info'].values[0]
                    corpo_email = f"{nome_pessoa}, bom dia!\n\n" \
                          "Segue em anexo planilha com pedidos pendentes de sua aprovação.\n\n" \
                          "Em caso de erros ou arquivos danificados entre em contato com o time de contas a pagar ou envie um email para:\n" \
                          "espro.gabriel@webmotors.com.br\n\n" \
                          "Atenciosamente,\n" \
                          "Contas a Pagar"
                    
                    assunto_email = f"[Automático] Pedidos Pendentes de Aprovação - {datetime.now().strftime('%d-%m-%Y')}"
                    enviar_email(email_destinatario, assunto_email, corpo_email, caminho_arquivo)

                    print(f"E-mail enviado para {nome_pessoa}.")

                else:
                    print(f"Nome '{nome_pessoa}' não encontrado no banco de dados de e-mails.")

            else:
                print(f"Arquivo {arquivo_excel} não encontrado.")

        print("E-mails enviados com sucesso!")

    else:
        print(f"Diretório {diretorio_pedidos} não encontrado.")

EmailPedidos()