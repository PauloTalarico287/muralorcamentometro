import pandas as pd
import requests
from datetime import datetime, timedelta
import json
import gspread
from google.oauth2 import service_account
import io
import os

# === CONFIGURAÇÃO DO LINK DINÂMICO ===
ano_completo = datetime.now().strftime("%Y")  # Ex: 2025
ano_curto = datetime.now().strftime("%y")     # Ex: 25
mes = datetime.now().strftime("%m")           # Ex: 10

url = f"https://orcamento.sf.prefeitura.sp.gov.br/orcamento/uploads/{ano_completo}/basedadosexecucao_{mes}{ano_curto}.xlsx"
print(f"🔗 Tentando baixar: {url}")

response = requests.get(url)
if response.status_code != 200:
    # Se não existir, tenta o mês anterior
    mes_anterior = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%m")
    url = f"https://orcamento.sf.prefeitura.sp.gov.br/orcamento/uploads/{ano_completo}/basedadosexecucao_{mes_anterior}{ano_curto}.xlsx"
    print(f"⚠️ Arquivo do mês atual não encontrado. Tentando link alternativo: {url}")
    response = requests.get(url)

bytes_data = io.BytesIO(response.content)
df = pd.read_excel(bytes_data)

# Seleção das colunas principais
orc = df[['Ds_Orgao', 'Vl_Orcado_Ano', 'Vl_Orcado_Atualizado', 'Vl_Congelado',  'Vl_Descongelado', 'Vl_Liquidado']]
Gastos = orc.groupby('Ds_Orgao')
investimento = Gastos.sum().reset_index()
investimento.columns = ['Órgão', 'Valor previsto para 2025', 'Valor Orçado Atualizado', 'Valor Congelado', 'Valor Descongelado', 'Realizado']

# === AUTENTICAÇÃO GOOGLE SHEETS ===
credentials_json = os.environ.get("GOOGLE_SHEETS_CREDENTIALS")
credentials_info = json.loads(credentials_json)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=scope)
gc = gspread.authorize(credentials)

spreadsheet_key = os.environ.get("GOOGLE_SHEETS_SPREADSHEET_KEY")
planilha = gc.open_by_key(spreadsheet_key)

# === SUBPREFEITURAS ===
investimento_por_sub = investimento[investimento['Órgão'].str.contains('Subprefeitura')]
investimento_por_sub = investimento_por_sub.query('Órgão != "Secretaria Municipal das Subprefeituras"')

# Mapeamento de nomes
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
investimento_por_sub['Órgão'] = investimento_por_sub['Órgão'].replace(dict(zip(nomes_existentes, novos_nomes)))
investimento_por_sub['Executado (%)'] = investimento_por_sub['Realizado']/investimento_por_sub['Valor previsto para 2025']*100
investimento_por_sub.sort_values('Executado (%)', ascending=False, inplace=True)

guia_sub = planilha.worksheet("Subprefeituras")
data_sub = [investimento_por_sub.columns.tolist()] + investimento_por_sub.values.tolist()
guia_sub.clear()
guia_sub.update(data_sub, 2)

# === SECRETARIAS ===
investimento_por_sec = investimento[investimento['Órgão'].str.contains('Secretaria')]
investimento_por_sec['Executado (%)'] = investimento_por_sec['Realizado']/investimento_por_sec['Valor previsto para 2025']*100
investimento_por_sec.sort_values('Executado (%)', ascending=False, inplace=True)

guia_sec = planilha.worksheet("Secretarias")
data_sec = [investimento_por_sec.columns.tolist()] + investimento_por_sec.values.tolist()
guia_sec.clear()
guia_sec.update(data_sec, 2)

# === OUTROS ÓRGÃOS ===
investimento_por_outros = investimento[~investimento['Órgão'].str.contains('Subprefeitura|Secretaria')]
investimento_por_outros['Executado (%)'] = investimento_por_outros['Realizado']/investimento_por_outros['Valor previsto para 2025']*100
investimento_por_outros.sort_values('Executado (%)', ascending=False, inplace=True)

guia_outros = planilha.worksheet("Outros")
data_outros = [investimento_por_outros.columns.tolist()] + investimento_por_outros.values.tolist()
guia_outros.clear()
guia_outros.update(data_outros, 2)

# === TOTAL ===
total_por_coluna = investimento.sum()
geral = pd.DataFrame({
    'Categoria': ['Total'],
    'Valor previsto para 2025': [total_por_coluna['Valor previsto para 2025']],
    'Realizado': [total_por_coluna['Realizado']],
    'Executado (%)': [(total_por_coluna['Realizado']/total_por_coluna['Valor previsto para 2025'])*100]
})

guia_geral = planilha.worksheet("Geral")
guia_geral.update('A2', geral['Categoria'].tolist()[0])
guia_geral.update('B2', geral['Valor previsto para 2025'].tolist()[0])
guia_geral.update('C2', geral['Realizado'].tolist()[0])
guia_geral.update('D2', geral['Executado (%)'].tolist()[0])

print("✅ Planilha atualizada com sucesso!")
