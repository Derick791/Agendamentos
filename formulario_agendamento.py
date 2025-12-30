import time
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe
from gspread.exceptions import APIError, WorksheetNotFound

# ============================
# Configurações
# ============================
LIMITE_VAGAS = 4
# ⚙️ Ajuste conforme sua planilha no Google Drive
SHEET_NAME = "InscricoesTreinamentos"   # nome do arquivo (documento) do Google Sheets
WORKSHEET_TITLE = "Inscricoes"          # nome da aba dentro da planilha (sem acento)
# Se preferir abrir por ID da planilha, defina abaixo:
SHEET_ID = "1Mys9TFql3h-NxFWfFueHLTpb_njbfcCyKPocBvigM2s"
# Se quiser apontar para a aba específica do link, use o gid (ex.: 0)
WORKSHEET_GID = 0  # mude para o gid da aba que você está olhando no Google Sheets

# Datas de treinamento (mantidas como no código original)
DATAS_TREINAMENTO = {
    "B1 - Substituir Caçamba Recuperadora Tipo Ponte": {
        "ADM (09-16h)": [
            "2025-09-22","2025-09-29","2025-10-06","2025-10-13","2025-10-20","2025-10-27",
            "2025-11-03","2025-11-10","2025-11-17","2025-11-24","2025-12-01","2025-12-08",
            "2025-12-15","2025-12-22","2025-12-29"
        ],
        "Noite (19h-02h)": [
            "2025-09-23","2025-09-30","2025-10-07","2025-10-14","2025-10-21","2025-10-28",
            "2025-11-04","2025-11-11","2025-11-18","2025-11-25","2025-12-02","2025-12-09",
            "2025-12-16","2025-12-23","2025-12-30"
        ],
    },
    "B2 - Substituir Cavaletes de Impacto articulado e rolos na mesa de impacto": {
        "ADM (09-16h)": [
            "2025-09-25","2025-10-02","2025-10-09","2025-10-16","2025-10-23","2025-10-30",
            "2025-11-06","2025-11-13","2025-11-20","2025-11-27","2025-12-04","2025-12-11",
            "2025-12-18","2025-12-25"
        ],
        "Noite (19h-02h)": [
            "2025-09-26","2025-10-03","2025-10-10","2025-10-17","2025-10-24","2025-10-31",
            "2025-11-07","2025-11-14","2025-11-21","2025-11-28","2025-12-05","2025-12-12",
            "2025-12-19","2025-12-26"
        ],
    },
    "B3 - Regular Freios Eletromagnéticos Do Giro da Lança Da EP2091KS e RCs 2092KS": {
        "ADM (09-16h)": [
            "2025-09-23","2025-09-30","2025-10-07","2025-10-14","2025-10-21","2025-10-28",
            "2025-11-04","2025-11-11","2025-11-18","2025-11-25","2025-12-02","2025-12-09",
            "2025-12-16","2025-12-23","2025-12-30"
        ],
        "Noite (19h-02h)": [
            "2025-09-25","2025-10-02","2025-10-09","2025-10-16","2025-10-23","2025-10-30",
            "2025-11-06","2025-11-13","2025-11-20","2025-11-27","2025-12-04","2025-12-11",
            "2025-12-18","2025-12-25"
        ],
    },
    "B4 - Substituir Atuador de Freio Vulkan SH13": {
        "ADM (09-16h)": [
            "2025-09-26","2025-10-03","2025-10-10","2025-10-17","2025-10-24","2025-10-31",
            "2025-11-07","2025-11-14","2025-11-21","2025-11-28","2025-12-05","2025-12-12",
            "2025-12-19","2025-12-26"
        ],
        "Noite (19h-02h)": [
            "2025-09-24","2025-10-01","2025-10-08","2025-10-15","2025-10-22","2025-10-29",
            "2025-11-05","2025-11-12","2025-11-19","2025-11-26","2025-12-03","2025-12-10",
            "2025-12-17","2025-12-24","2025-12-31"
        ],
    },
    "B5 - Realizar Substituição De Chapas De Revestimentos Silos e Chutes": {
        "ADM (09-16h)": [
            "2025-09-24","2025-10-01","2025-10-08","2025-10-15","2025-10-22","2025-10-29",
            "2025-11-05","2025-11-12","2025-11-19","2025-11-26","2025-12-03","2025-12-10",
            "2025-12-17","2025-12-24","2025-12-31"
        ],
        "Noite (19h-02h)": [
            "2025-09-22","2025-09-29","2025-10-06","2025-10-13","2025-10-20","2025-10-27",
            "2025-11-03","2025-11-10","2025-11-17","2025-11-24","2025-12-01","2025-12-08",
            "2025-12-15","2025-12-22","2025-12-29"
        ],
    },
}

# ============================
# Autenticação Google Sheets (via st.secrets)
# ============================
@st.cache_resource
def get_client():
    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
    except KeyError as e:
        st.error("⚠️ 'gcp_service_account' não encontrado em st.secrets. Configure suas credenciais do Google.")
        st.stop()
    except Exception as e:
        st.error(f"Falha ao carregar credenciais: {e}")
        st.stop()
    try:
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Falha ao autorizar cliente gspread: {e}")
        st.stop()

# Header esperado
EXPECTED_HEADER = [
    "Empresa", "Nome", "Matrícula", "Equipe/Gerência",
    "Treinamento", "Data", "Horário", "Turno"
]

def ensure_header(ws):
    """Garante que a primeira linha da planilha contém o cabeçalho esperado."""
    try:
        values = ws.get_all_values()
        if not values:
            ws.append_row(EXPECTED_HEADER, value_input_option="RAW")
        else:
            first_row = values[0]
            # Se a planilha foi criada vazia, pode vir com menos colunas
            if first_row != EXPECTED_HEADER:
                # Sobrescreve o cabeçalho na primeira linha
                ws.update('A1:H1', [EXPECTED_HEADER])
    except Exception as e:
        st.warning(f"Não foi possível verificar/atualizar o cabeçalho: {e}")

@st.cache_resource
def get_ws():
    """Retorna a worksheet (aba) especificada. Tenta pelo GID, depois pelo título; cria se não existir."""
    client = get_client()
    # Abre por ID (preferível) ou por nome
    try:
        sh = client.open_by_key(SHEET_ID) if SHEET_ID else client.open(SHEET_NAME)
    except APIError as e:
        st.error("⚠️ Não foi possível abrir a planilha. Verifique se o ID está correto e se a service account tem acesso como Editor.")
        st.error(f"Detalhes: {e}")
        st.stop()
    except Exception as e:
        st.error(f"Falha ao abrir a planilha: {e}")
        st.stop()

    # Tenta pelo GID primeiro
    ws = None
    if WORKSHEET_GID is not None:
        try:
            ws = sh.get_worksheet_by_id(WORKSHEET_GID)
        except WorksheetNotFound:
            ws = None
        except Exception as e:
            st.warning(f"Não foi possível abrir a aba pelo gid {WORKSHEET_GID}: {e}")
            ws = None

    # Fallback: tenta pelo título
    if ws is None:
        try:
            ws = sh.worksheet(WORKSHEET_TITLE)
        except WorksheetNotFound:
            try:
                ws = sh.add_worksheet(title=WORKSHEET_TITLE, rows=1000, cols=8)
                ws.append_row(EXPECTED_HEADER, value_input_option="RAW")
            except Exception as e:
                st.error(f"Falha ao criar a aba '{WORKSHEET_TITLE}': {e}")
                st.stop()
        except Exception as e:
            st.error(f"Falha ao abrir a aba '{WORKSHEET_TITLE}': {e}")
            st.stop()

    ensure_header(ws)
    return ws

# ============================
# Funções de negócio
# ============================

def _normalize(s: str) -> str:
    return (s or "").strip().casefold()

@st.cache_data(show_spinner=False)
def _read_all_rows():
    """Lê todos os valores (incluindo cabeçalho) para operações rápidas."""
    ws = get_ws()
    try:
        return ws.get_all_values()
    except Exception as e:
        st.error(f"Falha ao ler valores da planilha: {e}")
        return []

@st.cache_data(show_spinner=False)
def carregar_inscricoes() -> pd.DataFrame:
    ws = get_ws()
    try:
        df = get_as_dataframe(ws, evaluate_formulas=True, header=0)
        # Limpa linhas totalmente vazias
        df = df.dropna(how="all")
    except Exception:
        # Fallback robusto: reconstrói DF a partir de get_all_values
        values = _read_all_rows()
        if not values:
            return pd.DataFrame(columns=EXPECTED_HEADER)
        df = pd.DataFrame(values[1:], columns=values[0])
        df = df.dropna(how="all")

    # Garante todas as colunas esperadas
    for c in EXPECTED_HEADER:
        if c not in df.columns:
            df[c] = ""
    return df[EXPECTED_HEADER]

@st.cache_data(show_spinner=False)
def vagas_disponiveis(data: str, horario: str) -> int:
    values = _read_all_rows()
    dados = values[1:] if len(values) > 1 else []
    usados = 0
    for row in dados:
        try:
            # Índices: 5 -> Data, 6 -> Horário
            if len(row) >= 7 and row[5] == data and row[6] == horario:
                usados += 1
        except Exception:
            continue
    return max(LIMITE_VAGAS - usados, 0)


def safe_append_row(ws, values, retries: int = 3, base_delay: float = 0.8) -> bool:
    """Grava a linha com retentativas exponenciais em caso de erros transitórios (rate limit, etc.)."""
    for i in range(retries):
        try:
            ws.append_row(values, value_input_option="USER_ENTERED")
            return True
        except APIError as e:
            # Rate limit / permissão / erro de API
            if i < retries - 1:
                time.sleep(base_delay * (2 ** i))
            else:
                st.error("⚠️ Falha ao gravar na planilha (API). Verifique permissões ou tente novamente mais tarde.")
                st.error(f"Detalhes: {e}")
                return False
        except Exception as e:
            if i < retries - 1:
                time.sleep(base_delay * (2 ** i))
            else:
                st.error(f"Falha ao gravar na planilha: {e}")
                return False


def salvar_inscricao(empresa, nome, matricula, equipe, treinamento, data, horario, turno):
    ws = get_ws()

    # Carrega dados atuais para checagens
    values = _read_all_rows()
    dados = values[1:] if len(values) > 1 else []

    # Duplicidade: Nome (case-insensitive) + Treinamento + Data
    n_nome = _normalize(nome)
    n_trein = _normalize(treinamento)
    n_data = _normalize(data)

    for row in dados:
        try:
            if len(row) >= 6 and _normalize(row[1]) == n_nome and _normalize(row[4]) == n_trein and _normalize(row[5]) == n_data:
                st.error(f"{nome} já está inscrito neste treinamento nesta data.")
                return False
        except Exception:
            continue

    if vagas_disponiveis(data, horario) <= 0:
        st.error(f"As vagas para {data} ({horario}) já se esgotaram.")
        return False

    # Gravação
    ok = safe_append_row(ws, [empresa, nome, matricula, equipe, treinamento, data, horario, turno])
    if ok:
        # Limpa caches de leitura e força recarregar UI
        st.cache_data.clear()
        return True
    return False

# ============================
# App Streamlit (UI)
# ============================
st.title("📌 Formulário de Treinamentos")

empresa = st.selectbox("Empresa", ["Vale", "Parceira"])

nome = st.text_input("Nome completo")

matricula = ""
if empresa == "Vale":
    matricula = st.text_input("Matrícula (8 dígitos)")
    if matricula and (not matricula.isdigit() or len(matricula) != 8):
        st.warning("A matrícula deve ter exatamente 8 dígitos numéricos.")

# Gerência ou Parceira
if empresa == "Vale":
    equipe = st.selectbox("Gerência", ["Gerência de Pátio", "Gerência de Usina"])
else:
    equipe = st.selectbox("Parceira", ["Usimig", "Plagecon", "NDT"])

# Treinamento -> Horário -> Data
treinamento = st.selectbox("Treinamento", list(DATAS_TREINAMENTO.keys()))
horarios_disponiveis = list(DATAS_TREINAMENTO[treinamento].keys())
horario = st.selectbox("Horário", horarios_disponiveis)
datas_disponiveis = DATAS_TREINAMENTO[treinamento][horario]
data = st.selectbox("Data", datas_disponiveis)
turno = st.selectbox("Turno", ["Turno A", "Turno B", "Turno C", "Turno D"])

# Mostrar vagas
if data and horario:
    try:
        disponiveis = vagas_disponiveis(data, horario)
        st.info(f"🧾 Vagas disponíveis para {data} ({horario}): {disponiveis}/{LIMITE_VAGAS}")
    except Exception as e:
        st.warning(f"Não foi possível calcular vagas disponíveis: {e}")

# Botão salvar
if st.button("Salvar inscrição"):
    if not (empresa and nome and equipe and treinamento and data and horario and turno):
        st.warning("Preencha todos os campos obrigatórios.")
    elif empresa == "Vale" and (not matricula or len(matricula) != 8 or not matricula.isdigit()):
        st.warning("Matrícula inválida para funcionários da Vale.")
    else:
        if salvar_inscricao(empresa, nome, matricula, equipe, treinamento, data, horario, turno):
            st.success("✅ Inscrição registrada com sucesso!")
            # Recarrega a página para atualizar os quadros
            st.experimental_rerun()

# Resumo
st.markdown("---")
st.subheader("📈 Resumo para o instrutor")

df = carregar_inscricoes()
if df.empty:
    st.info("Nenhuma inscrição registrada até o momento.")
else:
    contagem = (
        df.groupby(["Treinamento", "Data", "Horário"]).size().reset_index(name="Inscritos")
    )
    contagem["Vagas Restantes"] = LIMITE_VAGAS - contagem["Inscritos"]
    st.write("### 👥 Turmas e vagas")
    st.dataframe(contagem)

    st.write("### 📋 Lista completa de inscritos")
    st.dataframe(df.sort_values(["Treinamento", "Data", "Horário"]))

    st.markdown("---")
    # Exportar CSV (opcional)
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Baixar inscrições (CSV)", data=csv, file_name="inscricoes.csv", mime="text/csv")
