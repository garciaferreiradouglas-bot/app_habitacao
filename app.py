import streamlit as st
from streamlit_drawable_canvas import st_canvas
import sqlite3
import pandas as pd
from PIL import Image
import io
import base64
import hashlib
from datetime import datetime
import numpy as np
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Habituação de Atiradores",
    page_icon="🎯",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- BANCO DE DADOS (SQLite) ---
DB_NAME = "habituacoes.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS atiradores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            cr TEXT NOT NULL UNIQUE
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS habituacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_hora TEXT NOT NULL,
            nome_atirador TEXT NOT NULL,
            cr_atirador TEXT NOT NULL,
            sigma_atirador TEXT,
            tipo_arma TEXT NOT NULL,
            tipo_municao TEXT NOT NULL,
            qtd_municao INTEGER NOT NULL,
            assinatura_base64 TEXT NOT NULL,
            hash_integridade TEXT NOT NULL
        )
    ''')
    
    c.execute("PRAGMA table_info(habituacoes)")
    colunas = [coluna[1] for coluna in c.fetchall()]
    if "sigma_atirador" not in colunas:
        c.execute("ALTER TABLE habituacoes ADD COLUMN sigma_atirador TEXT")

    conn.commit()
    conn.close()

init_db()

def popular_atiradores_exemplo():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM atiradores")
    if c.fetchone()[0] == 0:
        exemplos = [
            ("JOAO DA SILVA", "123456-CR"),
            ("MARIA OLIVEIRA", "654321-CR"),
            ("CARLOS SOUZA", "987654-CR")
        ]
        c.executemany("INSERT INTO atiradores (nome, cr) VALUES (?, ?)", exemplos)
        conn.commit()
    conn.close()

popular_atiradores_exemplo()

def get_lista_atiradores():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT id, nome || ' - CR: ' || cr AS rotulo, nome, cr FROM atiradores ORDER BY nome", conn)
    conn.close()
    return df

def salvar_habituação(nome, cr, sigma, arma, municao, qtd, img_array):
    im = Image.fromarray(img_array.astype('uint8'))
    buffered = io.BytesIO()
    im.save(buffered, format="PNG")
    assinatura_b64 = base64.b64encode(buffered.getvalue()).decode()

    data_hora_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    payload_validacao = f"{data_hora_str}|{nome}|{cr}|{sigma}|{arma}|{municao}|{qtd}|{assinatura_b64[:50]}"
    hash_integridade = hashlib.sha256(payload_validacao.encode('utf-8')).hexdigest()

    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        INSERT INTO habituacoes 
        (data_hora, nome_atirador, cr_atirador, sigma_atirador, tipo_arma, tipo_municao, qtd_municao, assinatura_base64, hash_integridade)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (data_hora_str, nome, cr, sigma, arma, municao, qtd, assinatura_b64, hash_integridade))
    conn.commit()
    conn.close()

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
        ws.cell(row=row_idx, column=1, value=row.id)
        ws.cell(row=row_idx, column=2, value=row.data_hora)
        ws.cell(row=row_idx, column=3, value=row.nome_atirador)
        ws.cell(row=row_idx, column=4, value=row.cr_atirador)
        ws.cell(row=row_idx, column=5, value=row.sigma_atirador or "")
        ws.cell(row=row_idx, column=6, value=row.tipo_arma)
        ws.cell(row=row_idx, column=7, value=row.tipo_municao)
        ws.cell(row=row_idx, column=8, value=row.qtd_municao)
        ws.cell(row=row_idx, column=10, value=row.hash_integridade)

        ws.row_dimensions[row_idx].height = 50

        if row.assinatura_base64:
            try:
                img_bytes = base64.b64decode(row.assinatura_base64)
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
    
    if df_atiradores.empty:
        st.warning("Nenhum atirador cadastrado no banco de dados.")
    else:
        termo_busca = st.text_input("🔍 Digite o nome ou CR para filtrar:", placeholder="Ex: Carlos ou 9876...")
        
        if termo_busca:
            df_filtrado = df_atiradores[
                df_atiradores['rotulo'].str.contains(termo_busca, case=False, na=False)
            ]
        else:
            df_filtrado = df_atiradores

        lista_opcoes = df_filtrado['rotulo'].tolist()

        if not lista_opcoes:
            st.error("Nenhum atirador encontrado com esse nome/CR.")
            opcao_selecionada = None
        else:
            opcao_selecionada = st.selectbox(
                "Selecione seu Nome / CR:",
                options=lista_opcoes,
                index=0
            )

        if opcao_selecionada:
            row_atirador = df_filtrado[df_filtrado['rotulo'] == opcao_selecionada].iloc[0]
            nome_atirador = row_atirador['nome']
            cr_atirador = row_atirador['cr']

            sigma_atirador = st.text_input("Número do SIGMA:", placeholder="Informe o número do SIGMA")

            col1, col2 = st.columns(2)
            with col1:
                tipo_arma = st.selectbox("Tipo de Arma:", ["Pistola", "Revólver", "Carabina", "Fuzil", "Espingarda"])
            with col2:
                tipo_municao = st.selectbox("Tipo/Calibre de Munição:", [
                    "9mm Luger", ".40 S&W", ".380 ACP", ".38 SPL", 
                    ".45 ACP", ".22 LR", "5.56x45mm", "12 GA"
                ])

            qtd_municao = st.number_input("Quantidade de Munição Utilizada:", min_value=1, max_value=1000, value=50, step=10)

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
                    salvar_habituação(
                        nome_atirador,
                        cr_atirador,
                        sigma_atirador,
                        tipo_arma,
                        tipo_municao,
                        qtd_municao,
                        img_data
                    )
                    st.session_state["canvas_key"] += 1
                    st.success("Habituação registrada com sucesso!")
                    st.balloons()
                    st.rerun()
                else:
                    st.error("Por favor, faça a assinatura antes de salvar.")

else:
    st.title("📊 Painel Administrativo")
    st.write("Visualização de registros, assinaturas e exportação.")

    conn = sqlite3.connect(DB_NAME)
    df_hab = pd.read_sql_query("SELECT id, data_hora, nome_atirador, cr_atirador, sigma_atirador, tipo_arma, tipo_municao, qtd_municao, assinatura_base64, hash_integridade FROM habituacoes ORDER BY id DESC", conn)
    conn.close()

    if df_hab.empty:
        st.info("Nenhuma habituação registrada até o momento.")
    else:
        df_exibicao = df_hab.drop(columns=['assinatura_base64'])
        st.dataframe(df_exibicao, use_container_width=True)

        st.subheader("🔍 Visualizar Registros e Assinaturas")
        
        for idx, row in df_hab.iterrows():
            with st.expander(f"ID #{row['id']} - {row['nome_atirador']} ({row['data_hora']})"):
                col_info, col_img = st.columns([1, 1])
                
                with col_info:
                    st.markdown(f"**Atirador:** {row['nome_atirador']}")
                    st.markdown(f"**CR:** {row['cr_atirador']}")
                    st.markdown(f"**SIGMA:** {row['sigma_atirador'] or 'N/A'}")
                    st.markdown(f"**Arma/Calibre:** {row['tipo_arma']} - {row['tipo_municao']}")
                    st.markdown(f"**Qtd. Munição:** {row['qtd_municao']}")
                    st.markdown(f"**Hash SHA-256:** `{row['hash_integridade']}`")
                
                with col_img:
                    st.markdown("**Assinatura Capturada:**")
                    try:
                        img_bytes = base64.b64decode(row['assinatura_base64'])
                        st.image(img_bytes, width=280)
                    except Exception as e:
                        st.error("Não foi possível carregar a imagem da assinatura.")

        st.markdown("---")
        
        excel_data = gerar_excel_com_assinaturas(df_hab)
        
        st.download_button(
            label="📥 Baixar Planilha Completa com Assinaturas (Excel)",
            data=excel_data,
            file_name=f"habituacoes_com_assinaturas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
