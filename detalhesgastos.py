import os
import json
import numpy as np
from datetime import datetime
import pandas as pd
import requests
import gspread
from google.oauth2 import service_account
import urllib3

from datetime import datetime, timedelta

# ✅ SEMPRE BUSCAR O DADO MAIS RECENTE DE 2025
ano_completo = "2025"
ano_curto = "25"

# Começar tentando o mês atual
mes_tentativa = datetime.now().month

# Se já estamos em 2026, começar por dezembro/2025
if datetime.now().year > 2025:
    mes_tentativa = 12

url = None
response = None

# Tentar meses de trás para frente até encontrar um arquivo disponível
while mes_tentativa >= 1:
    mes = str(mes_tentativa).zfill(2)  # Formata com zero à esquerda (01, 02, etc)
    url = f"https://orcamento.sf.prefeitura.sp.gov.br/orcamento/uploads/{ano_completo}/basedadosexecucao_{mes}{ano_curto}.xlsx"
    print(f"🔗 Tentando baixar: {url}")
    
    try:
        response = requests.get(url, timeout=30)
        if response.status_code == 200:
            print(f"✅ Arquivo encontrado: {mes}/{ano_completo}")
            break
    except requests.exceptions.SSLError:
        print("⚠️ Certificado HTTPS expirado, tentando com verificação desativada...")
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        response = requests.get(url, timeout=30, verify=False)
        if response.status_code == 200:
            print(f"✅ Arquivo encontrado: {mes}/{ano_completo}")
            break
    
    print(f"⚠️ Arquivo de {mes}/{ano_completo} não encontrado. Tentando mês anterior...")
    mes_tentativa -= 1

# Se não encontrou nenhum arquivo de 2025
if not response or response.status_code != 200:
    print(f"❌ Nenhum arquivo de 2025 encontrado. Código HTTP: {response.status_code if response else 'N/A'}")
    exit(0)

# Salvar o arquivo baixado
excel_file = f"basedadosexecucao_{mes}{ano_curto}.xlsx"
with open(excel_file, "wb") as f:
    f.write(response.content)

print(f"✅ Arquivo de {mes}/{ano_completo} baixado e salvo com sucesso!")

# ===============================
# 2️⃣ Ler e preparar dados
# ===============================
# ===============================
# 2️⃣ Ler e preparar dados
# ===============================
df = pd.read_excel(excel_file)

# Seleciona colunas relevantes
colunas = [
    'Ds_Orgao', 'Ds_Projeto_Atividade', 'Ds_Programa', 'Ds_Despesa',  # ✅ ADICIONADA
    'Vl_Orcado_Ano', 'Vl_Orcado_Atualizado',
    'Vl_Congelado', 'Vl_Descongelado', 'Vl_Liquidado'
]
df = df[colunas].copy()

# Renomear colunas
df.columns = [
    "Órgão", "Projeto/Atividade", "Programa", "Despesa",  # ✅ ADICIONADA
    "Previsto 2025", "Orçado Atualizado",
    "Congelado", "Descongelado", "Realizado"
]

# Substitui valores ausentes e zeros
df = df.replace([np.nan, np.inf, -np.inf], 0)

# Substitui valores ausentes e zeros
df = df.replace([np.nan, np.inf, -np.inf], 0)

# ✅ AGRUPAR E SOMAR por Órgão, Projeto/Atividade, Programa e Despesa
df = df.groupby(
    ["Órgão", "Projeto/Atividade", "Programa", "Despesa"], 
    as_index=False
).agg({
    "Previsto 2025": "sum",
    "Orçado Atualizado": "sum",
    "Congelado": "sum",
    "Descongelado": "sum",
    "Realizado": "sum"
})

# Calcula percentual executado
df["Executado (%)"] = np.where(
    df["Previsto 2025"] > 0,
    (df["Realizado"] / df["Previsto 2025"]) * 100,
    0
)

# Calcula percentual executado
df["Executado (%)"] = np.where(
    df["Previsto 2025"] > 0,
    (df["Realizado"] / df["Previsto 2025"]) * 100,
    0
)

# ===============================
# 3️⃣ Separar Subprefeituras e Outros Órgãos
# ===============================
subprefeituras = df[df["Órgão"].str.contains("Subprefeitura", case=False, na=False)].copy()
outros_orgaos = df[~df["Órgão"].str.contains("Subprefeitura", case=False, na=False)].copy()

subprefeituras = subprefeituras.sort_values(["Órgão", "Projeto/Atividade"])
outros_orgaos = outros_orgaos.sort_values(["Órgão", "Projeto/Atividade"])

print(f"🔹 Subprefeituras encontradas: {subprefeituras['Órgão'].nunique()}")
print(f"🔹 Outros órgãos encontrados: {outros_orgaos['Órgão'].nunique()}")

# ===============================
# 4️⃣ Autenticar Google Sheets
# ===============================
credentials_json = os.environ.get("GOOGLE_SHEETS_CREDENTIALS")
credentials_info = json.loads(credentials_json)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=scope)
gc = gspread.authorize(credentials)

spreadsheet_key = os.environ.get("GOOGLE_SHEETS_SPREADSHEET_KEY_DETALHES")
if not spreadsheet_key:
    raise Exception("❌ GOOGLE_SHEETS_SPREADSHEET_KEY_DETALHES não definido")

planilha = gc.open_by_key(spreadsheet_key)

# ===============================
# 5️⃣ Atualizar aba: Detalhes_Subprefeituras
# ===============================
try:
    guia_sub = planilha.worksheet("Detalhes_Subprefeituras")
except gspread.exceptions.WorksheetNotFound:
    guia_sub = planilha.add_worksheet(title="Detalhes_Subprefeituras", rows=5000, cols=20)

guia_sub.clear()

# Formatar valores numéricos
subprefeituras_formatted = subprefeituras.copy()
for col in ["Previsto 2025", "Orçado Atualizado", "Congelado", "Descongelado", "Realizado"]:
    subprefeituras_formatted[col] = subprefeituras_formatted[col].map(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

data_to_append_sub = [subprefeituras_formatted.columns.tolist()] + subprefeituras_formatted.values.tolist()
guia_sub.update(data_to_append_sub)
print("✅ Aba 'Detalhes_Subprefeituras' atualizada com sucesso!")

# ===============================
# 6️⃣ Atualizar aba: Detalhes_Outros_Orgaos
# ===============================
try:
    guia_outros = planilha.worksheet("Detalhes_Outros_Orgaos")
except gspread.exceptions.WorksheetNotFound:
    guia_outros = planilha.add_worksheet(title="Detalhes_Outros_Orgaos", rows=5000, cols=20)

guia_outros.clear()

outros_orgaos_formatted = outros_orgaos.copy()
for col in ["Previsto 2025", "Orçado Atualizado", "Congelado", "Descongelado", "Realizado"]:
    outros_orgaos_formatted[col] = outros_orgaos_formatted[col].map(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

data_to_append_outros = [outros_orgaos_formatted.columns.tolist()] + outros_orgaos_formatted.values.tolist()
guia_outros.update(data_to_append_outros)
print("✅ Aba 'Detalhes_Outros_Orgaos' atualizada com sucesso!")

# ===============================
# 7️⃣ Atualizar aba: Resumo por Programa (opcional)
# ===============================
programas = df.groupby(["Órgão", "Programa"], as_index=False).agg({
    "Previsto 2025": "sum",
    "Orçado Atualizado": "sum",
    "Congelado": "sum",
    "Descongelado": "sum",
    "Realizado": "sum"
})
programas["Executado (%)"] = np.where(
    programas["Previsto 2025"] > 0,
    (programas["Realizado"] / programas["Previsto 2025"]) * 100,
    0
)
programas = programas.replace([np.nan, np.inf, -np.inf], 0)

try:
    guia_prog = planilha.worksheet("Resumo_Programas")
except gspread.exceptions.WorksheetNotFound:
    guia_prog = planilha.add_worksheet(title="Resumo_Programas", rows=5000, cols=20)

guia_prog.clear()

programas_formatted = programas.copy()
for col in ["Previsto 2025", "Orçado Atualizado", "Congelado", "Descongelado", "Realizado"]:
    programas_formatted[col] = programas_formatted[col].map(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

data_to_append_prog = [programas_formatted.columns.tolist()] + programas_formatted.values.tolist()
guia_prog.update(data_to_append_prog)
print("✅ Aba 'Resumo_Programas' atualizada com sucesso!")

print("🎯 Processo concluído sem erros.")
