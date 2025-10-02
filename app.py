import json
import pandas as pd
import requests
from bs4 import BeautifulSoup
import gspread
import gspread
import os
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2 import service_account
from gspread_dataframe import get_as_dataframe, set_with_dataframe

try:
    url = "https://orcamento.sf.prefeitura.sp.gov.br/orcamento/uploads/2025/basedadosexecucao_1025.xlsx"
    response = requests.get(url)
    
    # Verifique se a solicitação foi bem-sucedida
    if response.status_code == 200:
        # Salve o conteúdo do arquivo em um arquivo local
        with open("basedadosexecucao1025.xlsx", "wb") as f:
            f.write(response.content)
    
        # Leia o arquivo Excel usando o pandas
        df = pd.read_excel("basedadosexecucao1025.xlsx")
    
        # Agora você pode trabalhar com os dados em 'df'
        print(df.head())
    
    else:
        print("Erro ao baixar o arquivo:", response.status_code)
    
    orcamento = pd.read_excel("basedadosexecucao1025.xlsx")
    orc=orcamento[['Ds_Orgao', 'Vl_Orcado_Ano', 'Vl_Orcado_Atualizado', 'Vl_Congelado',  'Vl_Descongelado', 'Vl_Liquidado']]
    Gastos=orc.groupby('Ds_Orgao')
    investimento=Gastos.sum()  
    investimento.sort_values('Vl_Liquidado', ascending=False)
    investimento = investimento.reset_index()
    novos = ['Órgão', 'Valor previsto para 2025', 'Valor Orçado Atualizado', 'Valor Congelado', 'Valor Descongelado', 'Realizado']
    investimento.columns = novos
    credentials_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS')
    if not credentials_json:
        raise Exception("❌ GOOGLE_SHEETS_CREDENTIALS não definido")

    #Corrigir quebras de linha e carregar JSON
    credentials_info = json.loads(credentials_json)
    
    # Definir escopo e criar credenciais
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    #credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=scope)
    credentials = service_account.Credentials.from_service_account_info(
        credentials_info,
        scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )

   
    # Autenticar gspread
    gc = gspread.authorize(credentials)
    
    spreadsheet_key = os.getenv('GOOGLE_SHEETS_SPREADSHEET_KEY')
    if not spreadsheet_key:
        raise Exception("GOOGLE_SHEETS_SPREADSHEET_KEY não definido")
    
    #SUBPREFEITURAS
    investimento_por_sub=investimento[investimento['Órgão'].str.contains('Subprefeitura')]
    investimento_por_sub = investimento_por_sub.query('Órgão != "Secretaria Municipal das Subprefeituras"')
    pd.set_option('float_format', '{:.2f}'.format)
    investimento_por_sub
    nomes_existentes = [
        "Subprefeitura Aricanduva/Formosa/Carrão",
        "Subprefeitura Butantã",
        "Subprefeitura Campo Limpo",
        "Subprefeitura Capela do Socorro",
        "Subprefeitura Casa Verde/Limão/Cachoeirinha",
        "Subprefeitura Cidade Ademar",
        "Subprefeitura Cidade Tiradentes",
        "Subprefeitura Ermelino Matarazzo",
        "Subprefeitura Freguesia/Brasilândia",
        "Subprefeitura Ipiranga",
        "Subprefeitura Itaim Paulista",
        "Subprefeitura Itaquera",
        "Subprefeitura Jabaquara",
        "Subprefeitura Jaçanã/Tremembé",
        "Subprefeitura Lapa",
        "Subprefeitura M'Boi Mirim",
        "Subprefeitura Mooca",
        "Subprefeitura Parelheiros",
        "Subprefeitura Penha",
        "Subprefeitura Perus/Anhanguera",
        "Subprefeitura Pinheiros",
        "Subprefeitura Pirituba/Jaraguá",
        "Subprefeitura Santana/Tucuruvi",
        "Subprefeitura Santo Amaro",
        "Subprefeitura Sapopemba",
        "Subprefeitura São Mateus",
        "Subprefeitura São Miguel Paulista",
        "Subprefeitura Sé",
        "Subprefeitura Vila Maria/Vila Guilherme",
        "Subprefeitura Vila Mariana",
        "Subprefeitura de Guaianases",
        "Subprefeitura de Vila Prudente"
    ]
    novos_nomes = [
        "Aricanduva/Vila Formosa",
        "Butantã",
        "Campo Limpo",
        "Capela do Socorro",
        "Casa Verde",
        "Cidade Ademar",
        "Cidade Tiradentes",
        "Ermelino Matarazzo",
        "Freguesia do Ó/Brasilândia",
        "Ipiranga",
        "Itaim Paulista",
        "Itaquera",
        "Jabaquara",
        "Jaçanã/Tremembé",
        "Lapa",
        "M'Boi Mirim",
        "Mooca",
        "Parelheiros",
        "Penha",
        "Perus",
        "Pinheiros",
        "Pirituba/Jaraguá",
        "Santana/Tucuruvi",
        "Santo Amaro",
        "Sapopemba",
        "São Mateus",
        "São Miguel",
        "Sé",
        "Vila Maria/Vila Guilherme",
        "Vila Mariana",
        "Guaianases",
        "Vila Prudente"
    ]
    mapeamento_nomes = dict(zip(nomes_existentes, novos_nomes))
    investimento_por_sub['Órgão'] = investimento_por_sub['Órgão'].replace(mapeamento_nomes)
    investimento_por_sub['Executado (%)'] = investimento_por_sub['Realizado']/investimento_por_sub['Valor previsto para 2025']*100
    investimento_por_sub.sort_values('Executado (%)', ascending=False)
    investimento_por_sub
    planilha = gc.open_by_key("1hURqGNnl9k4A_KhsQ4RvphmPGsJPlhq5NZpLnox0SUY")
    guia = planilha.worksheet("Subprefeituras")
    #data_to_append = investimento_por_sub.values.tolist()
    #guia.update(data_to_append)
    data_to_append = investimento_por_sub.values.tolist()
    data_to_append = [investimento_por_sub.columns.tolist()] + data_to_append
    
    guia.clear()
    guia.update(data_to_append, 2)
    #SECRETARIAS
    investimento_por_sec = investimento[investimento['Órgão'].str.contains('Secretaria')] 
    pd.set_option('float_format', '{:.2f}'.format)
    investimento_por_sec['Executado (%)'] = investimento_por_sec['Realizado']/investimento_por_sec['Valor previsto para 2025']*100
    investimento_por_sec.sort_values('Executado (%)', ascending=False)
    planilha = gc.open_by_key("1hURqGNnl9k4A_KhsQ4RvphmPGsJPlhq5NZpLnox0SUY")
    guia2 = planilha.worksheet("Secretarias")
    data_to_append2 = investimento_por_sec.values.tolist()
    data_to_append2 = [investimento_por_sec.columns.tolist()] + data_to_append2
    
    guia2.clear()
    guia2.update(data_to_append2, 2)
    
    #OUTROS_ORGAOS
    investimento_por_outros=investimento[~investimento['Órgão'].str.contains('Subprefeitura|Secretaria')]
    pd.set_option('float_format', '{:.2f}'.format)
    investimento_por_outros['Executado (%)'] = investimento_por_outros['Realizado']/investimento_por_outros['Valor previsto para 2025']*100
    investimento_por_outros.sort_values('Executado (%)', ascending=False)
    investimento_por_outros
    planilha = gc.open_by_key("1hURqGNnl9k4A_KhsQ4RvphmPGsJPlhq5NZpLnox0SUY")
    guia3 = planilha.worksheet("Outros")
    #data_to_append = investimento_por_sub.values.tolist()
    #guia.update(data_to_append)
    data_to_append3 = investimento_por_outros.values.tolist()
    data_to_append3 = [investimento_por_outros.columns.tolist()] + data_to_append3
    
    guia3.clear()
    guia3.update(data_to_append3, 2)
    
    #TOTAL
    total_por_coluna = investimento.sum()
    pd.set_option('float_format', '{:.2f}'.format)
    geral = pd.DataFrame({
        'Categoria': ['Total'],
        'Valor previsto para 2025': total_por_coluna['Valor previsto para 2025'],
        'Realizado': total_por_coluna['Realizado'],
        'Executado (%)': [(total_por_coluna['Realizado'] / total_por_coluna['Valor previsto para 2025']) * 100],
    })
    
    planilha = gc.open_by_key("1hURqGNnl9k4A_KhsQ4RvphmPGsJPlhq5NZpLnox0SUY")
    guia2 = planilha.worksheet("Geral")
    
    # Atualizando a fórmula de execução na célula correspondente
    linha_inicial = 2  # Pode ser ajustada conforme necessário
    #guia2.update_acell('D1', '=(C{} / B{}) * 100'.format(linha_inicial, linha_inicial))
    
    # Atualizando as células com os valores mais recentes
    guia2.update('A2', geral['Categoria'].tolist()[0])  # Acessando o primeiro elemento da lista
    guia2.update('B2', geral['Valor previsto para 2025'].tolist()[0])
    guia2.update('C2', geral['Realizado'].tolist()[0])
    guia2.update('D2', geral['Executado (%)'].tolist()[0])
    
except Exception as e:
    print("❌ Erro:", e)
    import traceback
    traceback.print_exc()
    exit(1)
