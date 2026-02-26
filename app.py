import pandas as pd
import requests
from datetime import datetime
import json
import gspread
from google.oauth2 import service_account
import io
import os
import urllib3

# === AUTENTICAÇÃO GOOGLE SHEETS ===
credentials_json = os.environ.get("GOOGLE_SHEETS_CREDENTIALS")
credentials_info = json.loads(credentials_json)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=scope)
gc = gspread.authorize(credentials)

spreadsheet_key = os.environ.get("GOOGLE_SHEETS_SPREADSHEET_KEY")
planilha = gc.open_by_key(spreadsheet_key)

# === FUNÇÃO AUXILIAR: obter ou criar aba ===
def get_or_create_worksheet(planilha, nome_aba):
    try:
        return planilha.worksheet(nome_aba)
    except gspread.exceptions.WorksheetNotFound:
        print(f"➕ Aba '{nome_aba}' não encontrada. Criando...")
        return planilha.add_worksheet(title=nome_aba, rows=200, cols=20)

# Mapeamento de nomes das subprefeituras
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
mapa_nomes = dict(zip(nomes_existentes, novos_nomes))

# === LOOP PELOS ANOS ===
for ano_atual in [2025, 2026]:
    ano_completo = str(ano_atual)
    ano_curto = str(ano_atual)[-2:]

    print(f"\n{'='*50}")
    print(f"📅 Processando dados de {ano_completo}...")
    print(f"{'='*50}")

    # Para 2025, começa em dezembro. Para outros anos, começa no mês atual
    if ano_atual < datetime.now().year:
        mes_tentativa = 12
    else:
        mes_tentativa = datetime.now().month

    response = None

    # Tentar meses de trás para frente até encontrar um arquivo disponível
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

    # Se não encontrou nenhum arquivo do ano
    if not response or response.status_code != 200:
        print(f"❌ Nenhum arquivo de {ano_completo} encontrado. Pulando para o próximo ano...")
        continue

    # Salvar e processar o arquivo
    excel_file = f"basedadosexecucao_{mes}{ano_curto}.xlsx"
    with open(excel_file, "wb") as f:
        f.write(response.content)
    print(f"✅ Arquivo de {mes}/{ano_completo} baixado e salvo com sucesso!")

    bytes_data = io.BytesIO(response.content)
    df = pd.read_excel(bytes_data)

    # Seleção das colunas principais
    orc = df[['Ds_Orgao', 'Vl_Orcado_Ano', 'Vl_Orcado_Atualizado', 'Vl_Congelado', 'Vl_Descongelado', 'Vl_Liquidado']]
    Gastos = orc.groupby('Ds_Orgao')
    investimento = Gastos.sum().reset_index()
    investimento.columns = ['Órgão', f'Valor previsto para {ano_completo}', 'Valor Orçado Atualizado', 'Valor Congelado', 'Valor Descongelado', 'Realizado']

    # Nome das abas: 2025 mantém nomes originais, anos seguintes ganham sufixo
    sufixo = "" if ano_atual == 2025 else f" {ano_completo}"

    # === SUBPREFEITURAS ===
    investimento_por_sub = investimento[investimento['Órgão'].str.contains('Subprefeitura')].copy()
    investimento_por_sub = investimento_por_sub.query('Órgão != "Secretaria Municipal das Subprefeituras"')
    investimento_por_sub['Órgão'] = investimento_por_sub['Órgão'].replace(mapa_nomes)
    investimento_por_sub['Executado (%)'] = investimento_por_sub['Realizado'] / investimento_por_sub[f'Valor previsto para {ano_completo}'] * 100
    investimento_por_sub.sort_values('Executado (%)', ascending=False, inplace=True)

    guia_sub = get_or_create_worksheet(planilha, f"Subprefeituras{sufixo}")
    data_sub = [investimento_por_sub.columns.tolist()] + investimento_por_sub.astype(object).values.tolist()
    guia_sub.clear()
    guia_sub.update(data_sub, 2)
    print(f"✅ Aba 'Subprefeituras{sufixo}' atualizada!")

    # === SECRETARIAS ===
    investimento_por_sec = investimento[investimento['Órgão'].str.contains('Secretaria')].copy()
    investimento_por_sec['Executado (%)'] = investimento_por_sec['Realizado'] / investimento_por_sec[f'Valor previsto para {ano_completo}'] * 100
    investimento_por_sec.sort_values('Executado (%)', ascending=False, inplace=True)

    guia_sec = get_or_create_worksheet(planilha, f"Secretarias{sufixo}")
    data_sec = [investimento_por_sec.columns.tolist()] + investimento_por_sec.astype(object).values.tolist()
    guia_sec.clear()
    guia_sec.update(data_sec, 2)
    print(f"✅ Aba 'Secretarias{sufixo}' atualizada!")

    # === OUTROS ÓRGÃOS ===
    investimento_por_outros = investimento[~investimento['Órgão'].str.contains('Subprefeitura|Secretaria')].copy()
    investimento_por_outros['Executado (%)'] = investimento_por_outros['Realizado'] / investimento_por_outros[f'Valor previsto para {ano_completo}'] * 100
    investimento_por_outros.sort_values('Executado (%)', ascending=False, inplace=True)

    guia_outros = get_or_create_worksheet(planilha, f"Outros{sufixo}")
    data_outros = [investimento_por_outros.columns.tolist()] + investimento_por_outros.astype(object).values.tolist()
    guia_outros.clear()
    guia_outros.update(data_outros, 2)
    print(f"✅ Aba 'Outros{sufixo}' atualizada!")

    # === TOTAL GERAL ===
    total_por_coluna = investimento.sum()
    guia_geral = get_or_create_worksheet(planilha, f"Geral{sufixo}")
    guia_geral.update('A2', 'Total')
    guia_geral.update('B2', float(total_por_coluna[f'Valor previsto para {ano_completo}']))
    guia_geral.update('C2', float(total_por_coluna['Realizado']))
    guia_geral.update('D2', float((total_por_coluna['Realizado'] / total_por_coluna[f'Valor previsto para {ano_completo}']) * 100))
    print(f"✅ Aba 'Geral{sufixo}' atualizada!")

print(f"\n🎉 Todos os anos processados com sucesso!")
