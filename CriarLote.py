from Imports import *

def CriarLote(file_path):
    df = pd.read_excel(file_path)

    df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
    df['Data de Vencimento'] = pd.to_datetime(df['Data de Vencimento'], format='%d/%m/%Y', errors='coerce')
    
    pedidos_pendentes = df[df['Status'] == 'Aprovação do supervisor pendente']
    aprovadores_unicos = pedidos_pendentes['Próximo aprovador'].unique()

    diretorio_atual = Path.cwd()
    pasta_destino = diretorio_atual / f'Pedidos_Enviados_{datetime.now().strftime("%d%m%Y")}'
    pasta_destino.mkdir(parents=True, exist_ok=True)

    for aprovador in aprovadores_unicos:
        pedidos_aprovador = pedidos_pendentes[pedidos_pendentes['Próximo aprovador'] == aprovador].copy()

        pedidos_aprovador['Data'] = pedidos_aprovador['Data'].dt.strftime('%d/%m/%Y')
        pedidos_aprovador['Data de Vencimento'] = pedidos_aprovador['Data de Vencimento'].dt.strftime('%d/%m/%Y')

        # Adicione as seguintes linhas para remover as colunas 'Valor líquido' e 'Pendente'
        pedidos_aprovador = pedidos_aprovador.drop(['Valor liquido', 'Pendente'], axis=1)

        caminho_saida = pasta_destino / f'{aprovador}.xlsx'
        pedidos_aprovador.to_excel(caminho_saida, index=False)

    print(f"Processo concluído. Arquivos Excel separados foram criados para cada aprovador na pasta {pasta_destino}")
