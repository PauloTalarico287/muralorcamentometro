import os
import json
import numpy as np
import pandas as pd
import requests
import gspread
from google.oauth2 import service_account

# ===============================
# 1️⃣ Baixar o Excel mais recente
# ===============================
ano_completo = datetime.now().strftime("%Y")  # Ex: 2025
ano_curto = datetime.now().strftime("%y")     # Ex: 25
mes = datetime.now().strftime("%m")           # Ex: 10

url = f"https://orcamento.sf.prefeitura.sp.gov.br/orcamento/uploads/{ano_completo}/basedadosexecucao_{mes}{ano_curto}.xlsx"

response = requests.get(url)
if response.status_code != 200:
    raise Exception(f"Erro ao baixar o arquivo: {response.status_code}")

excel_file = f"basedadosexecucao_{mes}{ano_curto}.xlsx"
with open(excel_file, "wb") as f:
    f.write(response.content)

print("✅ Arquivo baixado com sucesso!")

# ===============================
# 2️⃣ Ler e preparar dados
# ===============================
df = pd.read_excel(excel_file)

# Seleciona colunas relevantes
colunas = [
    'Ds_Orgao', 'Ds_Projeto_Atividade', 'Ds_Programa',
    'Vl_Orcado_Ano', 'Vl_Orcado_Atualizado',
    'Vl_Congelado', 'Vl_Descongelado', 'Vl_Liquidado'
]
df = df[colunas].copy()

# Renomear para colunas legíveis
df.columns = [
    "Órgão", "Projeto/Atividade", "Programa",
    "Previsto 2025", "Orçado Atualizado",
    "Congelado", "Descongelado", "Realizado"
]

# Substitui valores ausentes e zeros
df = df.replace([np.nan, np.inf, -np.inf], 0)

# Calcula percentual executado (evita divisão por zero)
df["Executado (%)"] = np.where(
    df["Previsto 2025"] > 0,
    (df["Realizado"] / df["Previsto 2025"]) * 100,
    0
)

# ===============================
# 3️⃣ Filtrar Subprefeituras
# ===============================
subprefeituras = df[df["Órgão"].str.contains("Subprefeitura", case=False, na=False)].copy()
subprefeituras = subprefeituras.sort_values(["Órgão", "Projeto/Atividade"])

print(f"🔹 Subprefeituras encontradas: {subprefeituras['Órgão'].nunique()}")

# ===============================
# 4️⃣ Autenticar Google Sheets
# ===============================
# ⚠️ Use o mesmo arquivo JSON que você usou no app.py
credentials_json = os.environ.get("GOOGLE_SHEETS_CREDENTIALS")
credentials_info = json.loads(credentials_json)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=scope)
gc = gspread.authorize(credentials)

# ID da planilha (igual ao que você já usou)
spreadsheet_key = os.getenv('GOOGLE_SHEETS_SPREADSHEET_KEY_DETALHES')
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

# Limpar e atualizar com dados novos
guia_sub.clear()

# Converter floats para string formatada (para evitar erro de JSON)
subprefeituras_formatted = subprefeituras.copy()
for col in ["Previsto 2025", "Orçado Atualizado", "Congelado", "Descongelado", "Realizado"]:
    subprefeituras_formatted[col] = subprefeituras_formatted[col].map(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

data_to_append = [subprefeituras_formatted.columns.tolist()] + subprefeituras_formatted.values.tolist()
guia_sub.update(data_to_append)

print("✅ Aba 'Detalhes_Subprefeituras' atualizada com sucesso!")

# ===============================
# 6️⃣ Agrupar por Programa (sem Subprefeituras)
# ===============================
programas = df[~df["Órgão"].str.contains("Subprefeitura", case=False, na=False)].copy()
programas = programas.groupby(["Órgão", "Programa"], as_index=False).agg({
    "Previsto 2025": "sum",
    "Orçado Atualizado": "sum",
    "Congelado": "sum",
    "Descongelado": "sum",
    "Realizado": "sum"
})

# Calcula percentual executado
programas["Executado (%)"] = np.where(
    programas["Previsto 2025"] > 0,
    (programas["Realizado"] / programas["Previsto 2025"]) * 100,
    0
)

programas = programas.replace([np.nan, np.inf, -np.inf], 0)

# ===============================
# 7️⃣ Atualizar aba: Detalhes_Programas
# ===============================
try:
    guia_prog = planilha.worksheet("Detalhes_Programas")
except gspread.exceptions.WorksheetNotFound:
    guia_prog = planilha.add_worksheet(title="Detalhes_Programas", rows=5000, cols=20)

guia_prog.clear()

# Formatar valores numéricos para evitar erro de JSON
programas_formatted = programas.copy()
for col in ["Previsto 2025", "Orçado Atualizado", "Congelado", "Descongelado", "Realizado"]:
    programas_formatted[col] = programas_formatted[col].map(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

data_to_append2 = [programas_formatted.columns.tolist()] + programas_formatted.values.tolist()
guia_prog.update(data_to_append2)

print("✅ Aba 'Detalhes_Programas' atualizada com sucesso!")
print("🎯 Processo concluído sem erros.")
