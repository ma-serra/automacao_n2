import streamlit as st
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2 import service_account

# Carregar as credenciais do Streamlit secrets
creds_info = st.secrets["gcp_service_account"]

# Criar as credenciais a partir do secrets
creds = service_account.Credentials.from_service_account_info(creds_info, scopes=['https://www.googleapis.com/auth/spreadsheets'])

# Crie o serviço do Google Sheets
service = build('sheets', 'v4', credentials=creds)

# Função para adicionar dados
def add_data(new_values, date_values, os_values, city_values):
    # Leia os dados existentes nas colunas B, C, D e E
    result_b = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME_B).execute()
    existing_values_b = result_b.get('values', [])

    result_c = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME_C).execute()
    existing_values_c = result_c.get('values', [])

    result_d = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME_D).execute()
    existing_values_d = result_d.get('values', [])

    result_e = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME_E).execute()
    existing_values_e = result_e.get('values', [])

    # Calcule a próxima linha para cada coluna
    next_row_b = len(existing_values_b) + 1  # A primeira linha é a 1
    next_row_c = len(existing_values_c) + 1  # A primeira linha é a 1
    next_row_d = len(existing_values_d) + 1  # A primeira linha é a 1
    next_row_e = len(existing_values_e) + 1  # A primeira linha é a 1

    # Crie o intervalo para atualizar a coluna B (exemplo: B2, B3, etc.)
    update_range_b = f'Monitoramento!B{next_row_b}'

    # Atualize a planilha com novos dados na coluna B
    body_b = {
        'values': new_values
    }
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=update_range_b,
        valueInputOption='RAW',
        body=body_b
    ).execute()

    # Crie o intervalo para atualizar a coluna C (exemplo: C2, C3, etc.)
    update_range_c = f'Monitoramento!C{next_row_c}'

    # Atualize a planilha com novos dados na coluna C
    body_c = {
        'values': city_values
    }
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=update_range_c,
        valueInputOption='RAW',
        body=body_c
    ).execute()

    # Crie o intervalo para atualizar a coluna D (exemplo: D2, D3, etc.)
    update_range_d = f'Monitoramento!D{next_row_d}'

    # Atualize a planilha com novos dados na coluna D
    body_d = {
        'values': date_values
    }
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=update_range_d,
        valueInputOption='RAW',
        body=body_d
    ).execute()

    # Crie o intervalo para atualizar a coluna E (exemplo: E2, E3, etc.)
    update_range_e = f'Monitoramento!E{next_row_e}'

    # Atualize a planilha com novos dados na coluna E
    body_e = {
        'values': os_values
    }
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=update_range_e,
        valueInputOption='RAW',
        body=body_e
    ).execute()

# Configuração da página
st.set_page_config(page_icon='content/file.png', page_title='N2', layout='wide')

# Solicitando arquivo
with st.container(border=True):
    dados = st.file_uploader('Anexe o arquivo XLS', type='xls')
    id_planilha = st.text_input('Cole o ID da planilha que deseja:')
    st.subheader('',divider='rainbow')
    
    # Passando informações e imagens necessárias
    st.markdown(':orange-background[Observação:] A planilha do Google Sheets e o formato do Dataframe que acrescentará os dados devem estar neste formato:')
    
    # Gerando a coluna para imagens exemplo
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(':orange-background[Exemplo Google Sheets:]')
        st.image('content/imagem_planilha.png')
    with col2:
        st.markdown(':orange-background[Exemplo Dataframe:]')
        st.image('content/imagem_dataframe.png', use_column_width=True)
        
    # Checando a quantidade de caracteres do ID da planilha
    quantidade_id = len(id_planilha)

    # Só passa para a próxima etapa quando atender os requisitos de tamanho e arquivo
    if (dados is not None) and (quantidade_id > 40):
        
        # Defina suas credenciais e informações da planilha
        SPREADSHEET_ID = id_planilha
        RANGE_NAME_B = 'Monitoramento!B1:B'  # Range para ler até a última linha da coluna B
        RANGE_NAME_D = 'Monitoramento!D1:D'  # Range para ler até a última linha da coluna D
        RANGE_NAME_E = 'Monitoramento!E1:E'  # Range para ler até a última linha da coluna E
        RANGE_NAME_C = 'Monitoramento!C1:C'  # Range para ler até a última linha da coluna C
        
        # Gerando o restante do front-end
        st.subheader('', divider='rainbow')
        dados_xls = pd.read_excel(dados, engine='xlrd')
        st.dataframe(dados_xls, use_container_width=True)  

        # Verifique se as colunas 'nome_cliente', 'data_abertura', 'tem_os' e 'cidade' existem
        if 'nome_cliente' in dados_xls.columns and 'data_abertura' in dados_xls.columns and 'tem_os' in dados_xls.columns and 'cidade' in dados_xls.columns:
            if st.button('Adicionar Dados'):
                # Capitalize a coluna 'tem_os'
                dados_xls['tem_os'] = dados_xls['tem_os'].dropna().apply(lambda x: x.capitalize())

                # Mapear as cidades para os códigos correspondentes
                city_mapping = {
                    'Ilhéus': 'ILH',
                    'Vitória da Conquista': 'VCA',
                    'Licínio de Almeida': 'LIC',
                    'Itabuna': 'ITB',
                    'Guanambi': 'GBI',
                    'Caculé': 'CLE',
                    'Luís Eduardo Magalhães': 'LEM',
                    'Caetité': 'CTE',
                    'Brumado': 'BRU',
                    'Barreiras': 'BRS',
                    'Jequié': 'JQE'
                }

                # Substituir os nomes das cidades pelos códigos
                dados_xls['cidade'].replace(city_mapping, inplace=True)

                # Extrair as colunas como listas de listas
                new_data = dados_xls['nome_cliente'].dropna().apply(lambda x: [x]).tolist()
                date_data = dados_xls['data_abertura'].dropna().apply(lambda x: [x]).tolist()
                os_data = dados_xls['tem_os'].dropna().apply(lambda x: [x]).tolist()
                city_data = dados_xls['cidade'].dropna().apply(lambda x: [x]).tolist()  # Nova linha para a cidade

                # Chame a função para adicionar dados
                add_data(new_data, date_data, os_data, city_data)
                st.success("Dados adicionados com sucesso!")  # Mensagem de sucesso sem negrito
        else:
            st.error("As colunas 'nome_cliente', 'data_abertura', 'tem_os' ou 'cidade' não foram encontradas no arquivo.")  # Mensagem de erro sem negrito
