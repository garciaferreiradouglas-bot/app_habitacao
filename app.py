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
    """Autentica lendo o JSON de credenciais guardado nas Secrets."""
    raw_json = st.secrets["gcp_service_account"]["json_data"]
    credentials_info = json.loads(raw_json)
    creds = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
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
    ws.append(headers)

    ws.column_dimensions['I'].width = 25

    for row_idx, row in enumerate(df_hab.itertuples(), start=2):
        ws.cell(row=row_idx, column=1, value=getattr(row, 'id', ''))
        ws.cell(row=row_idx, column=2, value=getattr(row, 'data_hora', ''))
        ws.cell(row=row_idx, column=3, value=getattr(row, 'nome_atirador', ''))
        ws.cell(row=row_idx, column=4, value=getattr(row, 'cr_atirador', ''))
        ws.cell(row=row_idx, column=5, value=getattr(row, 'sigma_atirador', ''))
        ws.cell(row=row_idx, column=6, value=getattr(row, 'tipo_arma', ''))
        ws.cell(row=row_idx, column=7, value=getattr(row, 'tipo_municao', ''))
        ws.cell(row=row_idx, column=8, value=getattr(row, 'qtd_municao', ''))
        ws.cell(row=row_idx, column=10, value=getattr(row, 'hash_integridade', ''))

        ws.row_dimensions[row_idx].height = 50

        assinatura_b64 = getattr(row, 'assinatura_base64', '')
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

# --- INICIALIZAÇÃO DO ESTADO ---
if "canvas_key" not in st.session_state:
    st.session_state["canvas_key"] = 0

# --- NAVEGAÇÃO ---
aba = st.radio("Selecione o Modo:", ["🎯 Registro de Habituação (Atirador)", "📊 Painel Admin / Exportar"], horizontal=True)

if aba == "🎯 Registro de Habituação (Atirador)":
    st.title("🎯 Registro de Habituação")
    st.write("Preencha os dados da sessão de tiro e assine no campo abaixo.")

    df_atiradores = get_lista_atiradores()
    
    termo_busca = st.text_input("🔍 Digite o nome ou CR para filtrar:", placeholder="Ex: Carlos ou 9876...")
    
    if termo_busca:
        df_filtrado = df_atiradores[
            df_atiradores['rotulo'].str.contains(termo_busca, case=False, na=False)
        ]
    else:
        df_filtrado = df_atiradores

    lista_opcoes = df_filtrado['rotulo'].tolist()

    if not lista_opcoes:
        st.warning("Nenhum atirador pré-cadastrado encontrado com esse termo. Preencha manualmente abaixo:")
        nome_atirador = st.text_input("Nome do Atirador:")
        cr_atirador = st.text_input("CR do Atirador:")
    else:
        opcao_selecionada = st.selectbox(
            "Selecione seu Nome / CR:",
            options=lista_opcoes,
            index=0
        )
        row_atirador = df_filtrado[df_filtrado['rotulo'] == opcao_selecionada].iloc[0]
        nome_atirador = row_atirador['nome']
        cr_atirador = row_atirador['cr']

    if nome_atirador and cr_atirador:
        sigma_atirador = st.text_input("Número do SIGMA:", placeholder="Informe o número do SIGMA")

        col1, col2 = st.columns(2)
        with col1:
            tipo_arma = st.selectbox("Tipo de Arma:", list(OPCOES_MUNICAO_POR_ARMA.keys()))
        with col2:
            municoes_disponiveis = OPCOES_MUNICAO_POR_ARMA[tipo_arma]
            tipo_municao = st.selectbox("Tipo/Calibre de Munição:", municoes_disponiveis)

        qtd_input = st.text_input("Quantidade de Munição Utilizada:", value="50", placeholder="Ex: 50, 100, 250...")

        st.subheader("🖋️ Assinatura Digital")
        st.caption("Assine dentro da caixa abaixo:")

        canvas_result = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=3,
            stroke_color="#000000",
            background_color="#F0F2F6",
            height=200,
            drawing_mode="freedraw",
            update_streamlit=True,
            return_image_data=True,
            key=f"canvas_assinatura_{st.session_state['canvas_key']}"
        )

        col_btn1, col_btn2 = st.columns([1, 1])
        with col_btn1:
            if st.button("🧹 Limpar Assinatura", use_container_width=True):
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
                st.error("Por favor, informe um número válido para a quantidade de munição.")
                st.stop()

            img_data = canvas_result.image_data if canvas_result is not None else None

            assinatura_valida = False
            if img_data is not None and isinstance(img_data, np.ndarray):
                if img_data.shape[2] == 4:
                    alpha = img_data[:, :, 3]
                    if np.any(alpha > 0):
                        assinatura_valida = True
                elif np.any(img_data < 250):
                    assinatura_valida = True

            if assinatura_valida:
                with st.spinner("Gravando no Google Sheets..."):
                    try:
                        salvar_habituação_sheets(
                            nome_atirador,
                            cr_atirador,
                            sigma_atirador,
                            tipo_arma,
                            tipo_municao,
                            qtd_municao,
                            img_data
                        )
                        st.session_state["canvas_key"] += 1
                        
                        st.markdown("""
                            <div style="text-align: center; padding: 20px; background-color: #d4edda; border-radius: 10px; border: 2px solid #28a745;">
                                <h1 style="color: #155724; margin: 0;">💥 🎯 💥</h1>
                                <h2 style="color: #155724; margin-top: 10px;">HABITUALIDADE CONCLUÍDA!</h2>
                                <p style="color: #155724; font-size: 18px;">Registro salvo com sucesso no Google Sheets.</p>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        time.sleep(5)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao salvar registro no Google Sheets: {e}")
            else:
                st.error("Por favor, faça a assinatura antes de salvar.")

else:
    st.title("📊 Painel Administrativo")
    st.write("Visualização de registros salvos no Google Sheets e exportação em Excel.")

    with st.spinner("Buscando registros da nuvem..."):
        df_hab = carregar_habituacoes_sheets()

    if df_hab.empty:
        st.info("Nenhuma habituação registrada até o momento no Google Sheets.")
    else:
        if 'assinatura_base64' in df_hab.columns:
            df_exibicao = df_hab.drop(columns=['assinatura_base64'])
        else:
            df_exibicao = df_hab

        st.dataframe(df_exibicao, use_container_width=True)

        st.subheader("🔍 Visualizar Registros e Assinaturas")
        
        for idx, row in df_hab.iterrows():
            nome = row.get('nome_atirador', 'N/A')
            dt = row.get('data_hora', 'N/A')
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
                    ass_b64 = row.get('assinatura_base64', '')
                    if ass_b64:
                        try:
                            img_bytes = base64.b64decode(ass_b64)
                            st.image(img_bytes, width=280)
                        except Exception:
                            st.error("Erro ao decodificar imagem da assinatura.")
                    else:
                        st.caption("Sem imagem de assinatura registrada.")

        st.markdown("---")
        
        excel_data = gerar_excel_com_assinaturas(df_hab)
        
        st.download_button(
            label="📥 Baixar Planilha Completa com Assinaturas (Excel)",
            data=excel_data,
            file_name=f"habituacoes_com_assinaturas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
