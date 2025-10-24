import os
import json
import numpy as np
from datetime import datetime
import pandas as pd
import requests
import gspread
from google.oauth2 import service_account
import io

# ===============================
# 1️⃣ Baixar o Excel mais recente
# ===============================
ano_completo = datetime.now().strftime("%Y")
ano_curto = datetime.now().strftime("%y")
mes = datetime.now().strftime("%m")

url = f"https://orcamento.sf.prefeitura.sp.gov.br/orcamento/uploads/{ano_completo}/basedadosexecucao_{mes}{ano_curto}.xlsx"
print(f"🔗 Baixando: {url}")

response = requests.get(url, timeout=30)
if response.status_code != 200:
    raise Exception(f"❌ Erro ao baixar o arquivo: {response.status_code}")

df = pd.read_excel(io.BytesIO(response.content))
print("✅ Arquivo baixado e lido com sucesso!")

# ===============================
# 2️⃣ Filtrar apenas Emendas Parlamentares
# ===============================
# Verificar se coluna existe
if 'TXT_VINC_PMSP' not in df.columns:
    raise Exception("❌ Coluna 'TXT_VINC_PMSP' não encontrada no arquivo")

# Filtrar linhas que contêm "Emendas Parlamentares"
emendas = df[df['TXT_VINC_PMSP'].str.contains('Emendas Parlamentares', case=False, na=False)].copy()

if emendas.empty:
    print("⚠️ Nenhuma emenda parlamentar encontrada!")
    exit()

print(f"📊 Total de registros de emendas: {len(emendas)}")

# ===============================
# 3️⃣ Dicionário de Partidos - LISTA OFICIAL (Legislatura 2025-2028)
# ===============================
PARTIDOS_VEREADORES = {
    # PT - 8 vereadores
    "Alessandro Guedes": "PT", "Dheison": "PT", "Dheison Silva": "PT",
    "Hélio Rodrigues": "PT", "Jair Tatto": "PT", "João Ananias": "PT",
    "Luna Zarattini": "PT", "Luana Zarattini": "PT", "Nabil Bonduki": "PT",
    "Senival Moura": "PT",

    # MDB - 7 vereadores
    "Ely Teruel": "MDB", "Fabio Riva": "MDB", "Fábio Riva": "MDB",
    "George Hato": "MDB", "João Jorge": "MDB", "Marcelo Messias": "MDB",
    "Paulo Frange": "MDB", "Sandra Santana": "MDB", "Sidney Cruz": "MDB", "Marlon Luz":"MDB",

    # PL - 7 vereadores
    "Dra Sandra Tadeu": "PL", "Dra. Sandra Tadeu": "PL",
    "Gilberto Nascimento": "PL", "Isac Félix": "PL", "Isac Felix": "PL",
    "Lucas Pavanato": "PL", "Rute Costa": "PL",
    "Sonaira Fernandes": "PL", "Zoe Martínez": "PL", "Zoe Martinez": "PL",
    "Missionário José Olimpio": "PL",

    # União Brasil - 7 vereadores
    "Adrilles Jorge": "União", "Amanda Vettorazzo": "União",
    "Pastora Sandra Alves": "União", "Ricardo Teixeira": "União",
    "Rubinho Nunes": "União", "Silvão Leite": "União",
    "Silvinho": "União", "Silvinho Leite": "União",

    # Podemos - 6 vereadores
    "Ana Carolina Oliveira": "Podemos", "Danilo do Posto": "Podemos",
    "Danilo do Posto de Saúde": "Podemos", "Dr. Milton Ferreira": "Podemos",
    "Gabriel Abreu": "Podemos", "Kenji Ito": "Podemos",
    "Kenji Palumbo": "Podemos", "Simone Ganem": "Podemos",

    # PSOL - 6 vereadores
    "Amanda Paschoal": "PSOL", "Celso Giannazi": "PSOL",
    "Keit Lima": "PSOL", "Luana Alves": "PSOL",
    "Professor Toninho Vespoli": "PSOL", "Toninho Vespoli": "PSOL",
    "Silvia da Bancada Feminista": "PSOL", "Silvia Ferraro": "PSOL",

    # PP - 4 vereadores
    "Dr. Murillo Lima": "PP", "Janaina Paschoal": "PP",
    "Major Palumbo": "PP", "Bombeiro Major Palumbo": "PP",
    "Sargento Nantes": "PP",

    # PSD - 3 vereadores
    "Edir Sales": "PSD", "Rodrigo Goulart": "PSD",
    "Thammy Miranda": "PSD", "Carlos Alberto Bezerra": "PSD",

    # Republicanos - 2 vereadores
    "André Santos": "Republicanos", "Sansão Pereira": "Republicanos", "André Souza":"Republicanos",

    # PSB - 2 vereadores
    "Eliseu Gabriel": "PSB", "Renata Falzoni": "PSB",

    # Novo - 1 vereador
    "Cris Monteiro": "Novo",

    # Rede - 1 vereador
    "Marina Bragante": "Rede",

    # PV - 1 vereador
    "Roberto Tripoli": "PV", "Tripoli": "PV",
}

# ===============================
# 4️⃣ Extrair nome do parlamentar e partido
# ===============================
# Exemplo: "Emendas Parlamentares - João Ananias" → "João Ananias"
emendas['Parlamentar'] = emendas['TXT_VINC_PMSP'].str.replace('Emendas Parlamentares - ', '', case=False, regex=False).str.strip()

# Filtrar emendas estaduais (não são vereadores municipais)
emendas = emendas[~emendas['Parlamentar'].str.contains('PMSP|Estadual', case=False, na=False)]

# Adicionar coluna de partido
emendas['Partido'] = emendas['Parlamentar'].map(PARTIDOS_VEREADORES)

# Alertar sobre parlamentares não mapeados
nao_mapeados = emendas[emendas['Partido'].isna()]['Parlamentar'].unique()
if len(nao_mapeados) > 0:
    print(f"⚠️ Parlamentares sem partido mapeado: {list(nao_mapeados)}")

# ===============================
# 5️⃣ Preparar DataFrame com colunas relevantes
# ===============================
colunas_necessarias = {
    'Ds_Orgao': 'Órgão',
    'Ds_Projeto_Atividade': 'Projeto/Atividade',
    'Ds_Programa': 'Programa',
    'Vl_Orcado_Atualizado': 'Orçado',
    'Vl_Liquidado': 'Liquidado',
    'Parlamentar': 'Parlamentar',
    'Partido': 'Partido'
}

# Verificar se todas as colunas existem (exceto Parlamentar e Partido que criamos)
colunas_faltantes = [col for col in colunas_necessarias.keys() 
                     if col not in ['Parlamentar', 'Partido'] and col not in df.columns]
if colunas_faltantes:
    print(f"⚠️ Colunas faltantes: {colunas_faltantes}")

emendas_clean = emendas[list(colunas_necessarias.keys())].copy()
emendas_clean.columns = list(colunas_necessarias.values())

# Substituir valores ausentes
emendas_clean = emendas_clean.replace([np.nan, np.inf, -np.inf], 0)

# Calcular % Executado
emendas_clean['Executado (%)'] = np.where(
    emendas_clean['Orçado'] > 0,
    (emendas_clean['Liquidado'] / emendas_clean['Orçado']) * 100,
    0
)

# ===============================
# 6️⃣ Criar Ranking por Parlamentar
# ===============================
ranking = emendas_clean.groupby(['Parlamentar', 'Partido'], as_index=False).agg({
    'Orçado': 'sum',
    'Liquidado': 'sum'
})

ranking['Executado (%)'] = np.where(
    ranking['Orçado'] > 0,
    (ranking['Liquidado'] / ranking['Orçado']) * 100,
    0
)

ranking['Qtd Projetos'] = emendas_clean.groupby(['Parlamentar', 'Partido']).size().values

# Ordenar por % Executado (maior para menor)
ranking = ranking.sort_values('Executado (%)', ascending=False)

# Adiciona coluna de posição no ranking
ranking.insert(0, 'Posição', range(1, len(ranking) + 1))

# Reordenar colunas
ranking = ranking[['Posição', 'Parlamentar', 'Partido', 'Qtd Projetos', 'Orçado', 'Liquidado', 'Executado (%)']]

print(f"\n🏆 Top 5 Parlamentares:")
print(ranking.head()[['Parlamentar', 'Partido', 'Liquidado', 'Executado (%)']].to_string(index=False))

# ===============================
# 7️⃣ Preparar Detalhamento
# ===============================
detalhamento = emendas_clean.sort_values(['Parlamentar', 'Liquidado'], ascending=[True, False])

# Reordenar colunas para mostrar Partido logo após Parlamentar
cols = detalhamento.columns.tolist()
cols.remove('Partido')
idx = cols.index('Parlamentar')
cols.insert(idx + 1, 'Partido')
detalhamento = detalhamento[cols]

# ===============================
# 8️⃣ Autenticar Google Sheets
# ===============================
# Tenta pegar das variáveis de ambiente (GitHub Actions)
credentials_json = os.environ.get("GOOGLE_SHEETS_CREDENTIALS")
credentials_info = json.loads(credentials_json)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=scope)
gc = gspread.authorize(credentials)

# Usar a planilha especificada (pode vir de variável de ambiente)
spreadsheet_key = os.environ.get("GOOGLE_SHEETS_SPREADSHEET_KEY_EMENDAS") or os.environ.get("GOOGLE_SHEETS_SPREADSHEET_KEY")
if not spreadsheet_key:
    raise Exception("❌ Nenhuma variável de planilha definida")
try:
    planilha = gc.open_by_key(spreadsheet_key)
    print(f"✅ Planilha encontrada: {planilha.title}")
except gspread.exceptions.SpreadsheetNotFound:
    raise Exception(f"❌ Planilha não encontrada! Verifique:\n"
                   f"1. Se o ID está correto: {spreadsheet_key}\n"
                   f"2. Se a conta de serviço tem acesso à planilha\n"
                   f"   Email: {credentials_info.get('client_email')}")

# ===============================
# 8️⃣ Atualizar aba: Ranking_Emendas
# ===============================
try:
    guia_ranking = planilha.worksheet("Ranking_Emendas")
except gspread.exceptions.WorksheetNotFound:
    guia_ranking = planilha.add_worksheet(title="Ranking_Emendas", rows=500, cols=10)

guia_ranking.clear()

# Formatar valores numéricos
ranking_formatted = ranking.copy()
for col in ['Orçado', 'Liquidado']:
    ranking_formatted[col] = ranking_formatted[col].map(
        lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )

data_ranking = [ranking_formatted.columns.tolist()] + ranking_formatted.values.tolist()
guia_ranking.update(data_ranking, 'A1')
print("✅ Aba 'Ranking_Emendas' atualizada!")

# ===============================
# 9️⃣ Atualizar aba: Detalhes_Emendas
# ===============================
try:
    guia_detalhes = planilha.worksheet("Detalhes_Emendas")
except gspread.exceptions.WorksheetNotFound:
    guia_detalhes = planilha.add_worksheet(title="Detalhes_Emendas", rows=5000, cols=10)

guia_detalhes.clear()

# Formatar valores numéricos
detalhamento_formatted = detalhamento.copy()
for col in ['Orçado', 'Liquidado']:
    detalhamento_formatted[col] = detalhamento_formatted[col].map(
        lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )

data_detalhes = [detalhamento_formatted.columns.tolist()] + detalhamento_formatted.values.tolist()
guia_detalhes.update(data_detalhes, 'A1')
print("✅ Aba 'Detalhes_Emendas' atualizada!")

print(f"\n🎯 Processo concluído!")
print(f"📊 Total de parlamentares: {len(ranking)}")
print(f"📋 Total de projetos: {len(detalhamento)}")
# Calcular totais ANTES da formatação (usar o ranking original)
total_orcado = ranking['Orçado'].sum() if ranking['Orçado'].dtype in ['float64', 'int64'] else emendas_clean['Orçado'].sum()
total_liquidado = ranking['Liquidado'].sum() if ranking['Liquidado'].dtype in ['float64', 'int64'] else emendas_clean['Liquidado'].sum()
print(f"💰 Total orçado: R$ {total_orcado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
print(f"✅ Total liquidado: R$ {total_liquidado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
