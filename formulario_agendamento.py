import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe

# ============================
# Configurações
# ============================
LIMITE_VAGAS = 4

# 👉 Ajuste conforme sua planilha no Google Drive
SHEET_NAME = "InscricoesTreinamentos"   # nome do arquivo (documento) do Google Sheets
WORKSHEET_TITLE = "Inscricoes"          # nome da aba dentro da planilha

# Se preferir abrir por ID da planilha, defina em Secrets: SHEET_ID = "..."
SHEET_ID = "1996jJ_zFo6dsQ7xzkWJY72lCZsefPurzIQH2aa7R9cw"

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
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return gspread.authorize(creds)

@st.cache_resource
def get_ws():
    """Retorna a worksheet (aba) 'Inscricoes'. Cria se não existir."""
    client = get_client()
    # Abre por ID (preferível) ou por nome
    if SHEET_ID:
        sh = client.open_by_key(SHEET_ID)
    else:
        sh = client.open(SHEET_NAME)
    try:
        ws = sh.worksheet(WORKSHEET_TITLE)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=WORKSHEET_TITLE, rows=1000, cols=8)
        ws.append_row([
            "Empresa", "Nome", "Matrícula", "Equipe/Gerência",
            "Treinamento", "Data", "Horário", "Turno"
        ])
    return ws

# ============================
# Funções de negócio
# ============================

def vagas_disponiveis(data: str, horario: str) -> int:
    ws = get_ws()
    valores = ws.get_all_values()
    dados = valores[1:] if len(valores) > 1 else []
    usados = sum(1 for row in dados if len(row) >= 7 and row[5] == data and row[6] == horario)
    return max(LIMITE_VAGAS - usados, 0)


def salvar_inscricao(empresa, nome, matricula, equipe, treinamento, data, horario, turno):
    ws = get_ws()
    valores = ws.get_all_values()
    dados = valores[1:] if len(valores) > 1 else []

    # Verificar duplicidade por (Nome, Treinamento, Data)
    for row in dados:
        if len(row) >= 6 and row[1] == nome and row[4] == treinamento and row[5] == data:
            st.error(f"{nome} já está inscrito neste treinamento nesta data.")
            return False

    if vagas_disponiveis(data, horario) <= 0:
        st.error(f"As vagas para {data} ({horario}) já se esgotaram.")
        return False

    ws.append_row([empresa, nome, matricula, equipe, treinamento, data, horario, turno])
    return True


def carregar_inscricoes() -> pd.DataFrame:
    ws = get_ws()
    df = get_as_dataframe(ws, evaluate_formulas=True, header=0)
    df = df.dropna(how="all")
    expected_cols = [
        "Empresa", "Nome", "Matrícula", "Equipe/Gerência",
        "Treinamento", "Data", "Horário", "Turno"
    ]
    for c in expected_cols:
        if c not in df.columns:
            df[c] = ""
    return df[expected_cols]

# ============================
# App Streamlit (UI)
# ============================

st.title("\U0001F4CC Formulário de Treinamentos")

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
    disponiveis = vagas_disponiveis(data, horario)
    st.info(f"\U0001F9EE Vagas disponíveis para {data} ({horario}): {disponiveis}/{LIMITE_VAGAS}")

# Botão salvar
if st.button("Salvar inscrição"):
    if not (empresa and nome and equipe and treinamento and data and horario and turno):
        st.warning("Preencha todos os campos obrigatórios.")
    elif empresa == "Vale" and (not matricula or len(matricula) != 8 or not matricula.isdigit()):
        st.warning("Matrícula inválida para funcionários da Vale.")
    else:
        if salvar_inscricao(empresa, nome, matricula, equipe, treinamento, data, horario, turno):
            st.success("\u2705 Inscrição registrada com sucesso!")

# Resumo
st.markdown("---")
st.subheader("\U0001F4C8 Resumo para o instrutor")
df = carregar_inscricoes()
if df.empty:
    st.info("Nenhuma inscrição registrada até o momento.")
else:
    contagem = (
        df.groupby(["Treinamento", "Data", "Horário"]).size().reset_index(name="Inscritos")
    )
    contagem["Vagas Restantes"] = LIMITE_VAGAS - contagem["Inscritos"]
    st.write("### \U0001F465 Turmas e vagas")
    st.dataframe(contagem)
    st.write("### \U0001F4CB Lista completa de inscritos")
    st.dataframe(df.sort_values(["Treinamento", "Data", "Horário"]))
    st.markdown("---")

# Exportar CSV (opcional)
csv = df.to_csv(index=False).encode("utf-8")
st.download_button("\u2B07\uFE0F Baixar inscrições (CSV)", data=csv, file_name="inscricoes.csv", mime="text/csv")
