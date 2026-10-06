import streamlit as st
from streamlit_drawable_canvas import st_canvas
import pandas as pd
from PIL import Image
import io
import base64
import hashlib
from datetime import datetime
import numpy as np
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage
import time
import uuid
import json
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Habituação de Atiradores",
    page_icon="🎯",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- AUTENTICAÇÃO E CONEXÃO COM GOOGLE SHEETS ---
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

@st.cache_resource
def get_gspread_client():
    """Autentica lendo os Secrets do Streamlit e tratando a chave privada RSA."""
    if "gcp_service_account" not in st.secrets:
        st.error("A seção [gcp_service_account] não foi encontrada nos Secrets do Streamlit.")
        st.stop()

    # Cria uma cópia mutável do dicionário obtido nos Secrets
    creds_info = dict(st.secrets["gcp_service_account"])

    # Se os dados vieram empacotados em um bloco 'json_data'
    if "json_data" in creds_info:
        raw_json = creds_info["json_data"]
        try:
            creds_info = json.loads(raw_json, strict=False)
        except Exception:
            raw_json_cleaned = raw_json.replace('\n', '\\n').replace('\r', '')
            creds_info = json.loads(raw_json_cleaned, strict=False)

    # Converte caracteres '\n' literais em quebras de linha reais exigidas pela chave RSA
    if "private_key" in creds_info:
        creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")

    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def get_sheet():
    """Obtém a folha principal do Google Sheets."""
    client = get_gspread_client()
    spreadsheet = client.open("Habituacoes_Clube")
    return spreadsheet.sheet1

# --- MAPEAMENTO DINÂMICO DE MUNIÇÕES POR TIPO DE ARMA ---
OPCOES_MUNICAO_POR_ARMA = {
    "Pistola": [
        "9mm Luger",
        ".380 ACP",
        ".40 S&W",
        ".45 ACP",
        ".38 TPC",
        ".22 LR",
        ".38 Super Auto",
        "10mm Auto",
        "7,65mm Browning",
        "6,35mm (.25 ACP)"
    ],
    "Revólver": [
        ".38 SPL",
        ".357 Magnum",
        ".22 LR",
        ".32 S&W / Longo",
        ".44 Magnum",
        ".454 Casull",
        ".22 WMR",
        ".44-40 WCF"
    ],
    "Carabina/Fuzil": [
        "5.56x45mm / .223 Rem",
        ".22 LR",
        "9mm Luger",
        ".300 Blackout",
        ".308 Win / 7.62x51mm",
        ".38 SPL",
        ".357 Magnum",
        ".40 S&W",
        ".380 ACP",
        ".17 HMR",
        "6.5 Creedmoor",
        ".44 Magnum",
        ".44-40 WCF",
        ".30-06 Springfield"
    ],
    "Espingarda": [
        "12 GA",
        "20 GA",
        "28 GA",
        "36 GA / .410",
        "16 GA",
        "24 GA",
        "32 GA"
    ]
}

# --- ATIRADORES CADASTRADOS ---
ATIRADORES_PADRAO = [
    {"nome": "JOAO DA SILVA", "cr": "123456-CR", "rotulo": "JOAO DA SILVA - CR: 123456-CR"},
    {"nome": "MARIA OLIVEIRA", "cr": "654321-CR", "rotulo": "MARIA OLIVEIRA - CR: 654321-CR"},
    {"nome": "CARLOS SOUZA", "cr": "987654-CR", "rotulo": "CARLOS SOUZA - CR: 987654-CR"}
]

def get_lista_atiradores():
    return pd.DataFrame(ATIRADORES_PADRAO)

# --- FUNÇÃO PARA SALVAR NO GOOGLE SHEETS ---
def salvar_habituação_sheets(nome, cr, sigma, arma, municao, qtd, img_array):
    im = Image.fromarray(img_array.astype('uint8'))
    buffered = io.BytesIO()
    im.save(buffered, format="PNG")
    assinatura_b64 = base64.b64encode(buffered.getvalue()).decode()

    data_hora_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    registro_id = str(uuid.uuid4())[:8]

    payload_validacao = f"{data_hora_str}|{nome}|{cr}|{sigma}|{arma}|{municao}|{qtd}|{assinatura_b64[:50]}"
    hash_integridade = hashlib.sha256(payload_validacao.encode('utf-8')).hexdigest()

    linha_dados = [
        registro_id,
        data_hora_str,
        nome,
        cr,
        sigma or "",
        arma,
        municao,
        int(qtd),
        assinatura_b64,
        hash_integridade
    ]

    sheet = get_sheet()
    sheet.append_row(linha_dados)

# --- FUNÇÃO PARA CARREGAR REGISTROS DO GOOGLE SHEETS ---
def carregar_habituacoes_sheets():
    try:
        sheet = get_sheet()
        registros = sheet.get_all_records()
        if not registros:
            return pd.DataFrame()
        df = pd.DataFrame(registros)
        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados do Google Sheets: {e}")
        return pd.DataFrame()

# --- FUNÇÃO PARA GERAR EXCEL COM ASSINATURAS ---
def gerar_excel_com_assinaturas(df_hab):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Habituações"

    headers = [
        "ID", "Data/Hora", "Nome Atirador", "CR", "SIGMA", 
        "Tipo Arma", "Tipo Munição", "Qtd. Munição", "Assinatura Digital", "Hash Integridade"
    ]
