import streamlit as st
from streamlit_drawable_canvas import st_canvas
import sqlite3
import pandas as pd
from PIL import Image
import io
import base64
import hashlib
from datetime import datetime

# --- CONFIGURAÇÃO DA PÁGINA (OTIMIZADA PARA TABLET) ---
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
    # Tabela de Atiradores
    c.execute('''
        CREATE TABLE IF NOT EXISTS atiradores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            cr TEXT NOT NULL UNIQUE
        )
    ''')
    # Tabela de Habituações
    c.execute('''
        CREATE TABLE IF NOT EXISTS habituacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_hora TEXT NOT NULL,
            nome_atirador TEXT NOT NULL,
            cr_atirador TEXT NOT NULL,
            tipo_arma TEXT NOT NULL,
            tipo_municao TEXT NOT NULL,
            qtd_municao INTEGER NOT NULL,
            assinatura_base64 TEXT NOT NULL,
            hash_integridade TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# --- CARGA INICIAL DE ATIRADORES DE EXEMPLO ---
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

# --- AUXILIARES DA APLICAÇÃO ---
def get_lista_atiradores():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT id, nome || ' - CR: ' || cr AS rotulo, nome, cr FROM atiradores ORDER BY nome", conn)
    conn.close()
    return df

def salvar_habituação(nome, cr, arma, municao, qtd, img_array):
    im = Image.fromarray(img_array.astype('uint8'))
    buffered = io.BytesIO()
    im.save(buffered, format="PNG")
    assinatura_b64 = base64.b64encode(buffered.getvalue()).decode()

    data_hora_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    payload_validacao = f"{data_hora_str}|{nome}|{cr}|{arma}|{municao}|{qtd}|{assinatura_b64[:50]}"
    hash_integridade = hashlib.sha256(payload_validacao.encode('utf-8')).hexdigest()

    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        INSERT INTO habituacoes 
        (data_hora, nome_atirador, cr_atirador, tipo_arma, tipo_municao, qtd_municao, assinatura_base64, hash_integridade)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (data_hora_str, nome, cr, arma, municao, qtd, assinatura_b64, hash_integridade))
    conn.commit()
    conn.close()

# --- NAVEGAÇÃO DA APLICAÇÃO ---
aba = st.radio("Selecione o Modo:", ["🎯 Registro de Habituação (Atirador)", "📊 Painel Admin / Exportar"], horizontal=True)

# ABA 1: REGISTRO
if aba == "🎯 Registro de Habituação (Atirador)":
    st.title("🎯 Registro de Habituação")
    st.write("Preencha os dados da sessão de tiro e assine no campo abaixo.")

    df_atiradores = get_lista_atiradores()
    
    if df_atiradores.empty:
        st.warning("Nenhum atirador cadastrado no banco de dados.")
    else:
        opcao_selecionada = st.selectbox(
            "Selecione seu Nome / CR:",
            options=df_atiradores['rotulo'].tolist(),
            index=0
        )
        
        row_atirador = df_atiradores[df_atiradores['rotulo'] == opcao_selecionada].iloc[0]
        nome_atirador = row_atirador['nome']
        cr_atirador = row_atirador['cr']

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
            width=500,
            drawing_mode="freedraw",
            key="canvas_assinatura",
        )

        st.info("📌 Registro com carimbo de tempo e hash criptográfico de validação (Lei 14.063/2020).")

        if st.button("✅ Registrar Habituação", type="primary", use_container_width=True):
            if canvas_result.image_data is not None and canvas_result.image_data.max() > 0:
                salvar_habituação(
                    nome_atirador,
                    cr_atirador,
                    tipo_arma,
                    tipo_municao,
                    qtd_municao,
                    canvas_result.image_data
                )
                st.success("Habituação registrada com sucesso!")
                st.balloons()
            else:
                st.error("Por favor, faça a assinatura antes de salvar.")

# ABA 2: ADMIN
else:
    st.title("📊 Painel Administrativo")
    st.write("Visualização de registros e exportação.")

    conn = sqlite3.connect(DB_NAME)
    df_hab = pd.read_sql_query("SELECT id, data_hora, nome_atirador, cr_atirador, tipo_arma, tipo_municao, qtd_municao, hash_integridade FROM habituacoes ORDER BY id DESC", conn)
    conn.close()

    if df_hab.empty:
        st.info("Nenhuma habituação registrada até o momento.")
    else:
        st.dataframe(df_hab, use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_hab.to_excel(writer, index=False, sheet_name='Habituações')
        
        st.download_button(
            label="📥 Baixar Planilha Completa (Excel)",
            data=buffer.getvalue(),
            file_name=f"habituacoes_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
