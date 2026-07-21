import os
import json
import numpy as np
from datetime import datetime
import pandas as pd
import requests
import gspread
from google.oauth2 import service_account
import urllib3
import io

# ===============================
# 1️⃣ Autenticar Google Sheets
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

# === FUNÇÃO AUXILIAR: obter ou criar aba ===
def get_or_create_worksheet(planilha, nome_aba, rows=5000):
    try:
        return planilha.worksheet(nome_aba)
    except gspread.exceptions.WorksheetNotFound:
        print(f"➕ Aba '{nome_aba}' não encontrada. Criando...")
        return planilha.add_worksheet(title=nome_aba, rows=rows, cols=20)

# === LOOP PELOS ANOS ===
for ano_atual in [2025, 2026]:
    ano_completo = str(ano_atual)
    ano_curto = str(ano_atual)[-2:]

    print(f"\n{'='*50}")
    print(f"📅 Processando dados de {ano_completo}...")
    print(f"{'='*50}")

    df = None

    # === 2026: URL fixa com CSV anual ===
    if ano_atual == 2026:
        url = "https://drive.prefeitura.sp.gov.br/cidade/secretarias/upload/seplan/arquivos/Exercicio_2026/basedadosexecucao_2026.csv"
        print(f"🔗 Baixando arquivo anual de 2026: {url}")
        try:
            response = requests.get(url, timeout=60)
            if response.status_code != 200:
                print(f"❌ Erro ao baixar arquivo de 2026: HTTP {response.status_code}")
                continue
        except requests.exceptions.SSLError:
            print("⚠️ Certificado HTTPS expirado, tentando com verificação desativada...")
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            try:
                response = requests.get(url, timeout=60, verify=False)
                if response.status_code != 200:
                    print(f"❌ Erro ao baixar arquivo de 2026: HTTP {response.status_code}")
                    continue
            except Exception as e:
                print(f"❌ Erro ao baixar arquivo de 2026: {e}")
                continue
        except Exception as e:
            print(f"❌ Erro ao baixar arquivo de 2026: {e}")
            continue

        with open("basedadosexecucao_2026.csv", "wb") as f:
            f.write(response.content)
        print(f"✅ Arquivo de 2026 baixado!")

        bytes_data = io.BytesIO(response.content)
        df = pd.read_csv(bytes_data, sep=';', encoding='latin1', decimal=',')

    # === 2025: busca mensal normal (xlsx) ===
    else:
        if ano_atual < datetime.now().year:
            mes_tentativa = 12
        else:
            mes_tentativa = datetime.now().month

        response = None

        while mes_tentativa >= 1:
            mes = str(mes_tentativa).zfill(2)
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

        if not response or response.status_code != 200:
            print(f"❌ Nenhum arquivo de {ano_completo} encontrado. Pulando para o próximo ano...")
            continue

        excel_file = f"basedadosexecucao_{mes}{ano_curto}.xlsx"
        with open(excel_file, "wb") as f:
            f.write(response.content)
        print(f"✅ Arquivo de {mes}/{ano_completo} baixado e salvo com sucesso!")

        bytes_data = io.BytesIO(response.content)
        df = pd.read_excel(bytes_data)

    if df is None:
        print(f"❌ Não foi possível carregar dados de {ano_completo}. Pulando...")
        continue

    print(f"✅ Arquivo carregado! Total de linhas: {len(df)}")

    # ===============================
    # 2️⃣ Ler e preparar dados
    # ===============================
    col_previsto = f"Previsto {ano_completo}"

    colunas = [
        'Ds_Orgao', 'Ds_Projeto_Atividade', 'Ds_Programa', 'Ds_Despesa',
        'Vl_Orcado_Ano', 'Vl_Orcado_Atualizado',
        'Vl_Congelado', 'Vl_Descongelado', 'Vl_Liquidado'
    ]
    df = df[colunas].copy()

    df.columns = [
        "Órgão", "Projeto/Atividade", "Programa", "Despesa",
        col_previsto, "Orçado Atualizado",
        "Congelado", "Descongelado", "Realizado"
    ]

    df = df.replace([np.nan, np.inf, -np.inf], 0)

    df = df.groupby(
        ["Órgão", "Projeto/Atividade", "Programa", "Despesa"],
        as_index=False
    ).agg({
        col_previsto: "sum",
        "Orçado Atualizado": "sum",
        "Congelado": "sum",
        "Descongelado": "sum",
        "Realizado": "sum"
    })

    df["Executado (%)"] = np.where(
        df[col_previsto] > 0,
        (df["Realizado"] / df[col_previsto]) * 100,
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

    # Nome das abas: 2025 sem sufixo, 2026 em diante com sufixo
    sufixo = "" if ano_atual == 2025 else f" {ano_completo}"

    # Função para formatar valores numéricos
    def formatar_numeros(df_in, col_previsto):
        df_fmt = df_in.copy()
        for col in [col_previsto, "Orçado Atualizado", "Congelado", "Descongelado", "Realizado"]:
            df_fmt[col] = df_fmt[col].map(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        return df_fmt

    # ===============================
    # 4️⃣ Atualizar aba: Detalhes_Subprefeituras
    # ===============================
    guia_sub = get_or_create_worksheet(planilha, f"Detalhes_Subprefeituras{sufixo}")
    guia_sub.clear()
    sub_fmt = formatar_numeros(subprefeituras, col_previsto)
    guia_sub.update([sub_fmt.columns.tolist()] + sub_fmt.values.tolist())
    print(f"✅ Aba 'Detalhes_Subprefeituras{sufixo}' atualizada!")

    # ===============================
    # 5️⃣ Atualizar aba: Detalhes_Outros_Orgaos
    # ===============================
    guia_outros = get_or_create_worksheet(planilha, f"Detalhes_Outros_Orgaos{sufixo}")
    guia_outros.clear()
    outros_fmt = formatar_numeros(outros_orgaos, col_previsto)
    guia_outros.update([outros_fmt.columns.tolist()] + outros_fmt.values.tolist())
    print(f"✅ Aba 'Detalhes_Outros_Orgaos{sufixo}' atualizada!")

    # ===============================
    # 6️⃣ Atualizar aba: Resumo_Programas
    # ===============================
    programas = df.groupby(["Órgão", "Programa"], as_index=False).agg({
        col_previsto: "sum",
        "Orçado Atualizado": "sum",
        "Congelado": "sum",
        "Descongelado": "sum",
        "Realizado": "sum"
    })
    programas["Executado (%)"] = np.where(
        programas[col_previsto] > 0,
        (programas["Realizado"] / programas[col_previsto]) * 100,
        0
    )
    programas = programas.replace([np.nan, np.inf, -np.inf], 0)

    guia_prog = get_or_create_worksheet(planilha, f"Resumo_Programas{sufixo}")
    guia_prog.clear()
    prog_fmt = formatar_numeros(programas, col_previsto)
    guia_prog.update([prog_fmt.columns.tolist()] + prog_fmt.values.tolist())
    print(f"✅ Aba 'Resumo_Programas{sufixo}' atualizada!")

print(f"\n🎯 Processo concluído para todos os anos!")
