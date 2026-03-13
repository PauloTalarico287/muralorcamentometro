import os
import json
import numpy as np
from datetime import datetime
import pandas as pd
import requests
import gspread
from google.oauth2 import service_account
import urllib3

# ===============================
# 1️⃣ Autenticar Google Sheets
# ===============================
credentials_json = os.environ.get("GOOGLE_SHEETS_CREDENTIALS")
credentials_info = json.loads(credentials_json)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=scope)
gc = gspread.authorize(credentials)

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

# === FUNÇÃO AUXILIAR: obter ou criar aba ===
def get_or_create_worksheet(planilha, nome_aba, rows=5000):
    try:
        return planilha.worksheet(nome_aba)
    except gspread.exceptions.WorksheetNotFound:
        print(f"➕ Aba '{nome_aba}' não encontrada. Criando...")
        return planilha.add_worksheet(title=nome_aba, rows=rows, cols=10)

# ===============================
# 2️⃣ Dicionário de Partidos
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
    "Paulo Frange": "MDB", "Sandra Santana": "MDB", "Sidney Cruz": "MDB", "Marlon Luz": "MDB",

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
    "Rodolfo Despachante": "União",

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
    "André Santos": "Republicanos", "Sansão Pereira": "Republicanos", "André Souza": "Republicanos",

    # PSB - 2 vereadores
    "Eliseu Gabriel": "PSB", "Renata Falzoni": "PSB",

    # Novo - 1 vereador
    "Cris Monteiro": "Novo",

    # Rede - 1 vereador
    "Marina Bragante": "Rede",

    # PV - 1 vereador
    "Roberto Tripoli": "PV", "Tripoli": "PV",
}

# === LOOP PELOS ANOS ===
for ano_atual in [2025, 2026]:
    ano_completo = str(ano_atual)
    ano_curto = str(ano_atual)[-2:]

    print(f"\n{'='*50}")
    print(f"📅 Processando emendas de {ano_completo}...")
    print(f"{'='*50}")

    # Para anos passados começa em dezembro, para o atual começa no mês corrente
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

    # Salvar e carregar arquivo
    excel_file = f"basedadosexecucao_{mes}{ano_curto}.xlsx"
    with open(excel_file, "wb") as f:
        f.write(response.content)
    print(f"✅ Arquivo de {mes}/{ano_completo} baixado e salvo com sucesso!")

    print(f"📖 Carregando arquivo Excel...")
    df = pd.read_excel(excel_file)
    print(f"✅ Arquivo carregado! Total de linhas: {len(df)}")

    # ===============================
    # 3️⃣ Filtrar apenas Emendas Parlamentares
    # ===============================
    if 'TXT_VINC_PMSP' not in df.columns:
        print(f"❌ Coluna 'TXT_VINC_PMSP' não encontrada para {ano_completo}. Pulando...")
        continue

    emendas = df[df['TXT_VINC_PMSP'].str.contains('Emendas Parlamentares', case=False, na=False)].copy()

    if emendas.empty:
        print(f"⚠️ Nenhuma emenda parlamentar encontrada em {ano_completo}!")
        continue

    print(f"📊 Total de registros de emendas: {len(emendas)}")

    # ===============================
    # 4️⃣ Extrair nome do parlamentar e partido
    # ===============================
    emendas['Parlamentar'] = emendas['TXT_VINC_PMSP'].str.replace('Emendas Parlamentares - ', '', case=False, regex=False).str.strip()
    emendas = emendas[~emendas['Parlamentar'].str.contains('PMSP|Estadual', case=False, na=False)]
    emendas['Partido'] = emendas['Parlamentar'].map(PARTIDOS_VEREADORES)

    nao_mapeados = emendas[emendas['Partido'].isna()]['Parlamentar'].unique()
    if len(nao_mapeados) > 0:
        print(f"⚠️ Parlamentares sem partido mapeado: {list(nao_mapeados)}")

    # ===============================
    # 5️⃣ Preparar DataFrame
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

    colunas_faltantes = [col for col in colunas_necessarias.keys()
                         if col not in ['Parlamentar', 'Partido'] and col not in df.columns]
    if colunas_faltantes:
        print(f"⚠️ Colunas faltantes: {colunas_faltantes}")

    emendas_clean = emendas[list(colunas_necessarias.keys())].copy()
    emendas_clean.columns = list(colunas_necessarias.values())
    emendas_clean = emendas_clean.replace([np.nan, np.inf, -np.inf], 0)

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
    ranking = ranking.sort_values('Executado (%)', ascending=False)
    ranking.insert(0, 'Posição', range(1, len(ranking) + 1))
    ranking = ranking[['Posição', 'Parlamentar', 'Partido', 'Qtd Projetos', 'Orçado', 'Liquidado', 'Executado (%)']]

    print(f"\n🏆 Top 5 Parlamentares:")
    print(ranking.head()[['Parlamentar', 'Partido', 'Liquidado', 'Executado (%)']].to_string(index=False))

    # ===============================
    # 7️⃣ Preparar Detalhamento
    # ===============================
    detalhamento = emendas_clean.sort_values(['Parlamentar', 'Liquidado'], ascending=[True, False])
    cols = detalhamento.columns.tolist()
    cols.remove('Partido')
    idx = cols.index('Parlamentar')
    cols.insert(idx + 1, 'Partido')
    detalhamento = detalhamento[cols]

    # Nome das abas: 2025 sem sufixo, 2026 em diante com sufixo
    sufixo = "" if ano_atual == 2025 else f" {ano_completo}"

    # Função de formatação
    def formatar_numeros(df_in, colunas):
        df_fmt = df_in.copy()
        for col in colunas:
            df_fmt[col] = df_fmt[col].map(
                lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                if isinstance(x, (int, float)) else x
            )
        return df_fmt

    # ===============================
    # 8️⃣ Atualizar aba: Ranking_Emendas
    # ===============================
    guia_ranking = get_or_create_worksheet(planilha, f"Ranking_Emendas{sufixo}", rows=500)
    guia_ranking.clear()
    ranking_fmt = formatar_numeros(ranking, ['Orçado', 'Liquidado'])
    guia_ranking.update([ranking_fmt.columns.tolist()] + ranking_fmt.values.tolist(), 'A1')
    print(f"✅ Aba 'Ranking_Emendas{sufixo}' atualizada!")

    # ===============================
    # 9️⃣ Atualizar aba: Detalhes_Emendas
    # ===============================
    guia_detalhes = get_or_create_worksheet(planilha, f"Detalhes_Emendas{sufixo}", rows=5000)
    guia_detalhes.clear()
    detalhamento_fmt = formatar_numeros(detalhamento, ['Orçado', 'Liquidado'])
    guia_detalhes.update([detalhamento_fmt.columns.tolist()] + detalhamento_fmt.values.tolist(), 'A1')
    print(f"✅ Aba 'Detalhes_Emendas{sufixo}' atualizada!")

    # Totais
    total_orcado = emendas_clean['Orçado'].sum()
    total_liquidado = emendas_clean['Liquidado'].sum()
    print(f"\n🎯 Processo concluído para {ano_completo}!")
    print(f"📊 Total de parlamentares: {len(ranking)}")
    print(f"📋 Total de projetos: {len(detalhamento)}")
    print(f"💰 Total orçado: R$ {total_orcado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    print(f"✅ Total liquidado: R$ {total_liquidado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

print(f"\n🎉 Todos os anos processados com sucesso!")
