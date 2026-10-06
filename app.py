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
    """Autentica lendo os Secrets em formato JSON nativo do Streamlit Cloud."""
    if "gcp_service_account" not in st.secrets:
        st.error("A seção [gcp_service_account] não foi encontrada nos Secrets do Streamlit.")
        st.stop()

    sec_data = st.secrets["gcp_service_account"]

    if "json_data" in sec_data:
        raw_json = sec_data["json_data"]
        if isinstance(raw_json, str):
            creds_info = json.loads(raw_json, strict=False)
        else:
            creds_info = dict(raw_json)
    else:
        creds_info = dict(sec_data)

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
        "9mm Luger", ".380 ACP", ".40 S&W", ".45 ACP", ".38 TPC", 
        ".22 LR", ".38 Super Auto", "10mm Auto", "7,65mm Browning", "6,35mm (.25 ACP)"
    ],
    "Revólver": [
        ".38 SPL", ".357 Magnum", ".22 LR", ".32 S&W / Longo", 
        ".44 Magnum", ".454 Casull", ".22 WMR", ".44-40 WCF"
    ],
    "Carabina/Fuzil": [
        "5.56x45mm / .223 Rem", ".22 LR", "9mm Luger", ".300 Blackout", 
        ".308 Win / 7.62x51mm", ".38 SPL", ".357 Magnum", ".40 S&W", 
        ".380 ACP", ".17 HMR", "6.5 Creedmoor", ".44 Magnum", ".44-40 WCF", ".30-06 Springfield"
    ],
    "Espingarda": [
        "12 GA", "20 GA", "28 GA", "36 GA / .410", "16 GA", "24 GA", "32 GA"
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

# --- CONVERTE MATRIZ DO CANVAS EM BASE64 ROBUSTO ---
def processar_imagem_canvas(img_array):
    try:
        img_uint8 = img_array.astype(np.uint8)
        
        # Cria imagem PIL a partir do array (RGBA)
        img_pil = Image.fromarray(img_uint8, mode="RGBA")
        
        # Cria fundo branco e combina com o desenho em preto
        background = Image.new("RGB", img_pil.size, (255, 255, 255))
        background.paste(img_pil, mask=img_pil.split()[3]) # canal Alpha como máscara
        
        buffered = io.BytesIO()
        background.save(buffered, format="PNG")
        return base64.b64encode(buffered.getvalue()).decode('utf-8')
    except Exception as e:
        st.error(f"Erro ao processar imagem da assinatura: {e}")
        return ""

# --- SALVA NO GOOGLE SHEETS ---
def salvar_habituação_sheets(nome, cr, sigma, arma, municao, qtd, img_array):
    assinatura_b64 = processar_imagem_canvas(img_array)
    
    if not assinatura_b64:
        raise ValueError("Não foi possível processar o desenho da assinatura.")

    data_hora_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    registro_id = str(uuid.uuid4())[:8]

    # Cálculo do Hash SHA-256
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

# --- CARREGA REGISTROS DA PLANILHA ---
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

# --- BUSCA A COLUNA DA ASSINATURA INDEPENDENTE DO NOME ---
def obter_valor_assinatura(row):
    """Busca a string Base64 em qualquer coluna de assinatura existente na linha."""
    # Nomes comuns de coluna
    colunas_possiveis = ['assinatura_base64a', 'assinatura_base64', 'assinatura', 'Assinatura Digital']
    
    for col in colunas_possiveis:
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != "":
            val_str = str(row[col]).strip()
            
            # Se contiver prefixos de URL/Fórmula
            if "base64," in val_str:
                try:
                    return val_str.split("base64,")[1].split('"')[0].split("'")[0].split(')')[0].strip()
                except Exception:
                    pass
            elif len(val_str) > 100 and not val_str.startswith("http") and not val_str.startswith("="):
                return val_str
                
    # Tenta pegar pela 9ª coluna (índice 8) caso a busca por nome falhe
    try:
        val_ind = row.iloc[8]
        if pd.notna(val_ind) and len(str(val_ind)) > 100:
            return str(val_ind).strip()
    except Exception:
        pass

    return ""

# --- EXPORTAÇÃO EXCEL COM IMAGENS DAS ASSINATURAS ---
def gerar_excel_com_assinaturas(df_hab):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Habituações"

    headers = [
        "ID", "Data/Hora", "Nome Atirador", "CR", "SIGMA", 
        "Tipo Arma", "Tipo Munição", "Qtd. Munição", "Assinatura Digital", "Hash Integridade"
    ]
    ws.append(headers)

    ws.column_dimensions['I'].width = 25

    for row_idx, (_, row) in enumerate(df_hab.iterrows(), start=2):
        ws.cell(row=row_idx, column=1, value=row.get('id', ''))
        ws.cell(row=row_idx, column=2, value=row.get('data_hora', ''))
        ws.cell(row=row_idx, column=3, value=row.get('nome_atirador', ''))
        ws.cell(row=row_idx, column=4, value=row.get('cr_atirador', ''))
        ws.cell(row=row_idx, column=5, value=row.get('sigma_atirador', ''))
        ws.cell(row=row_idx, column=6, value=row.get('tipo_arma', ''))
        ws.cell(row=row_idx, column=7, value=row.get('tipo_municao', ''))
        ws.cell(row=row_idx, column=8, value=row.get('qtd_municao', ''))
        ws.cell(row=row_idx, column=10, value=row.get('hash_integridade', ''))

        ws.row_dimensions[row_idx].height = 55

        assinatura_b64 = obter_valor_assinatura(row)

        if assinatura_b64:
            try:
                img_bytes = base64.b64decode(assinatura_b64)
                img_file = io.BytesIO(img_bytes)
                img = OpenPyxlImage(img_file)
                img.width = 140
                img.height = 60
                
                cell_address = f"I{row_idx}"
                ws.add_image(img, cell_address)
            except Exception:
                pass

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# --- ESTADO INICIAL ---
if "canvas_key" not in st.session_state:
    st.session_state["canvas_key"] = 0

# --- ABAS DE NAVEGAÇÃO ---
aba = st.radio("Selecione o Modo:", ["🎯 Registro de Habituação (Atirador)", "📊 Painel Admin / Exportar"], horizontal=True)

if aba == "🎯 Registro de Habituação (Atirador)":
    st.title("🎯 Registro de Habituação")
    st.write("Preencha os dados da sessão de tiro e assine no campo abaixo.")

    df_atiradores = get_lista_atiradores()
    termo_busca = st.text_input("🔍 Digite o nome ou CR para filtrar:", placeholder="Ex: Carlos ou 9876...")
    
    if termo_busca:
        df_filtrado = df_atiradores[df_atiradores['rotulo'].str.contains(termo_busca, case=False, na=False)]
    else:
        df_filtrado = df_atiradores

    lista_opcoes = df_filtrado['rotulo'].tolist()

    if not lista_opcoes:
        nome_atirador = st.text_input("Nome do Atirador:")
        cr_atirador = st.text_input("CR do Atirador:")
    else:
        opcao_selecionada = st.selectbox("Selecione seu Nome / CR:", options=lista_opcoes, index=0)
        row_atirador = df_filtrado[df_filtrado['rotulo'] == opcao_selecionada].iloc[0]
        nome_atirador = row_atirador['nome']
        cr_atirador = row_atirador['cr']

    if nome_atirador and cr_atirador:
        sigma_atirador = st.text_input("Número do SIGMA:", placeholder="Informe o número do SIGMA")

        col1, col2 = st.columns(2)
        with col1:
            tipo_arma = st.selectbox("Tipo de Arma:", list(OPCOES_MUNICAO_POR_ARMA.keys()))
        with col2:
            tipo_municao = st.selectbox("Tipo/Calibre de Munição:", OPCOES_MUNICAO_POR_ARMA[tipo_arma])

        qtd_input = st.text_input("Quantidade de Munição Utilizada:", value="50")

        st.subheader("🖋️ Assinatura Digital")
        st.caption("Assine dentro da caixa abaixo:")

        canvas_result = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=3,
            stroke_color="#000000",
            background_color="#FFFFFF",
            height=200,
            drawing_mode="freedraw",
            update_streamlit=True,
            return_image_data=True,
            key=f"canvas_assinatura_{st.session_state['canvas_key']}"
        )

        if st.button("🧹 Limpar Assinatura"):
            st.session_state["canvas_key"] += 1
            st.rerun()

        st.info("📌 Registro com carimbo de tempo e hash criptográfico de validação (Lei 14.063/2020).")

        if st.button("✅ Registrar Habituação", type="primary", use_container_width=True):
            try:
                qtd_municao = int(qtd_input.strip())
                if qtd_municao <= 0:
                    st.error("A quantidade de munição deve ser maior que zero.")
                    st.stop()
            except ValueError:
                st.error("Por favor, informe um número válido para a quantidade.")
                st.stop()

            img_data = canvas_result.image_data if canvas_result is not None else None

            # Validação do desenho
            assinatura_valida = False
            if img_data is not None and isinstance(img_data, np.ndarray):
                if img_data.shape[2] == 4:
                    alpha = img_data[:, :, 3]
                    if np.any(alpha > 0):
                        assinatura_valida = True

            if assinatura_valida:
                with st.spinner("Gravando no Google Sheets..."):
                    try:
                        salvar_habituação_sheets(
                            nome_atirador, cr_atirador, sigma_atirador,
                            tipo_arma, tipo_municao, qtd_municao, img_data
                        )
                        st.session_state["canvas_key"] += 1
                        st.success("HABITUALIDADE REGISTRADA COM SUCESSO!")
                        time.sleep(3)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")
            else:
                st.error("Por favor, assine o campo de assinatura antes de salvar.")

else:
    st.title("📊 Painel Administrativo")
    st.write("Visualização de registros salvos no Google Sheets e exportação em Excel.")

    with st.spinner("Buscando registros da nuvem..."):
        df_hab = carregar_habituacoes_sheets()

    if df_hab.empty:
        st.info("Nenhuma habituação registrada até o momento.")
    else:
        # Oculta colunas longas da tabela de visão geral
        cols_ocultar = [c for c in df_hab.columns if 'assinatura' in c.lower()]
        df_exibicao = df_hab.drop(columns=cols_ocultar, errors='ignore')
        st.dataframe(df_exibicao, use_container_width=True)

        st.subheader("🔍 Visualizar Registros e Assinaturas")
        
        for idx, row in df_hab.iterrows():
            nome = row.get('nome_atirador', '')
            dt = row.get('data_hora', '')
            reg_id = row.get('id', idx)
            
            with st.expander(f"ID #{reg_id} - {nome} ({dt})"):
                col_info, col_img = st.columns([1, 1])
                
                with col_info:
                    st.markdown(f"**Atirador:** {row.get('nome_atirador', '')}")
                    st.markdown(f"**CR:** {row.get('cr_atirador', '')}")
                    st.markdown(f"**SIGMA:** {row.get('sigma_atirador', 'N/A')}")
                    st.markdown(f"**Arma/Calibre:** {row.get('tipo_arma', '')} - {row.get('tipo_municao', '')}")
                    st.markdown(f"**Qtd. Munição:** {row.get('qtd_municao', '')}")
                    st.markdown(f"**Hash SHA-256:** `{row.get('hash_integridade', '')}`")
                
                with col_img:
                    st.markdown("**Assinatura Capturada:**")
                    ass_b64 = obter_valor_assinatura(row)
                    if ass_b64:
                        try:
                            img_bytes = base64.b64decode(ass_b64)
                            st.image(img_bytes, width=280)
                        except Exception:
                            st.error("Erro ao converter imagem Base64.")
                    else:
                        st.caption("Sem imagem de assinatura válida nesta linha.")

        st.markdown("---")
        
        excel_data = gerar_excel_com_assinaturas(df_hab)
        
        st.download_button(
            label="📥 Baixar Planilha Completa com Assinaturas (Excel)",
            data=excel_data,
            file_name=f"habituacoes_com_assinaturas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
