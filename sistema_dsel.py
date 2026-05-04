import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime
import os
import hashlib

# --- FUNÇÃO DE SEGURANÇA ---
def gerar_hash(texto):
    return hashlib.sha256(texto.encode()).hexdigest()

# --- 1. CONFIGURAÇÕES E ESTILO ---
st.set_page_config(page_title="DSEL - GESTÃO NOVACAP", layout="wide")
DB_NAME = 'estoque_dsel_v3.db'
NOME_LOGO = "logo.novacap.png"

# --- 2. BANCO DE DADOS ---
def conectar(): return sqlite3.connect(DB_NAME)

def inicializar_banco():
    with conectar() as conn:
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS usuarios (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                      nome TEXT UNIQUE, senha TEXT, status TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS estoque (id INTEGER PRIMARY KEY AUTOINCREMENT, item TEXT, cor TEXT, 
                      tamanho TEXT, modelo TEXT, qtd INTEGER, estado TEXT, nf TEXT, pessoa TEXT, 
                      data_hora TEXT, almoxarife TEXT)''')
        
        # Senha padrão: admin123 (Em formato Hash para o GitHub)
        senha_hash = gerar_hash("admin123")
        c.execute("INSERT OR IGNORE INTO usuarios (nome, senha, status) VALUES (?, ?, ?)", ('ADMIN', senha_hash, 'Ativo'))
        
        # --- CORREÇÃO AUTOMÁTICA DE MODELOS ---
        c.execute("UPDATE estoque SET modelo = 'BRIM' WHERE modelo = 'Calça BRIM'")
        c.execute("UPDATE estoque SET modelo = 'Padrão' WHERE modelo IN ('Manga Curta', 'Manga Longa')")
        conn.commit()

inicializar_banco()

# --- 3. FUNÇÕES DE APOIO ---
def exibir_logo(largura=200):
    if os.path.exists(NOME_LOGO): 
        try: st.image(NOME_LOGO, width=largura)
        except: pass

def consultar_saldo(item, cor, tamanho, modelo, estado):
    with conectar() as conn:
        res = conn.execute("""SELECT SUM(qtd) FROM estoque 
                            WHERE item=? AND cor=? AND tamanho=? AND modelo=? AND estado=?""", 
                           (item, cor, tamanho, modelo, estado)).fetchone()
        return res[0] if res and res[0] is not None else 0

# --- 4. CONTROLE DE ACESSO ---
if 'logado' not in st.session_state: st.session_state.logado = False

if not st.session_state.logado:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        exibir_logo(250)
        st.title("🔐 Acesso DSEL")
        u = st.text_input("Usuário").strip().upper()
        s = st.text_input("Senha", type="password")
        if st.button("Entrar", use_container_width=True):
            s_hash = gerar_hash(s)
            with conectar() as conn:
                res = conn.execute("SELECT nome FROM usuarios WHERE UPPER(nome)=? AND senha=? AND status='Ativo'", (u, s_hash)).fetchone()
            if res:
                st.session_state.logado, st.session_state.almoxarife = True, res[0]
                st.rerun()
            else: st.error("Acesso negado.")
    st.stop()

# --- LISTAS PADRONIZADAS ---
CORES = ["Cinza", "Amarelo", "Azul", "Branco", "Verde", "Preto", "Marrom", "Laranja", "Vermelho"]
ITENS = ["Camiseta", "Calça", "Bota", "Luvas", "Chapéu", "Boné", "Cinto", "Meias"]
TAMANHOS = ["Único","P", "M", "G", "GG", "XG","XGG"] + [str(i).zfill(2) for i in range(34, 57)]
MODELOS = ["Padrão", "Botina","Bontina Galocha","Luva de couro","luva vaqueta", "Cano Alto", "Refletivo","BRIM","JEANS","Elastano","Tactel","Camuflado"]

with st.sidebar:
    exibir_logo(150)
    st.subheader(f"👤 {st.session_state.almoxarife}")
    opcoes = ["🏠 Painel Principal", "📥 Entrada (Fábrica)", "📦 Entrega (Saída)", "📊 Relatório Geral"]
    menu = st.radio("Navegação", opcoes)
    if st.button("Sair"): st.session_state.logado = False; st.rerun()

# --- 5. MÓDULOS ---
if menu == "🏠 Painel Principal":
    st.header("🏠 Detalhamento de Itens em Estoque")
    item_sel = st.selectbox("Selecione um item:", ["-- Selecione --"] + ITENS)
    if item_sel != "-- Selecione --":
        with conectar() as conn:
            df_det = pd.read_sql_query("""SELECT modelo, cor, tamanho, estado, SUM(qtd) as 'Saldo' 
                                        FROM estoque WHERE item=? GROUP BY modelo, cor, tamanho, estado HAVING SUM(qtd) != 0""", conn, params=(item_sel,))
        st.subheader("📊 Saldo Atual")
        st.dataframe(df_det, use_container_width=True, hide_index=True)

elif menu == "📥 Entrada (Fábrica)":
    st.header("📥 Entrada da Fábrica")
    with st.form("ent"):
        nf = st.text_input("Nota Fiscal")
        it, co, ta, mod = st.selectbox("Item", ITENS), st.selectbox("Cor", CORES), st.selectbox("Tamanho", TAMANHOS), st.selectbox("Modelo", MODELOS)
        qt = st.number_input("Quantidade", min_value=1)
        if st.form_submit_button("Salvar Entrada"):
            with conectar() as conn:
                conn.execute("INSERT INTO estoque (item, cor, tamanho, modelo, qtd, estado, nf, pessoa, data_hora, almoxarife) VALUES (?,?,?,?,?,?,?,?,?,?)", (it, co, ta, mod, qt, "Novo", nf, "Fábrica", datetime.now().strftime("%d/%m/%Y %H:%M"), st.session_state.almoxarife))
            st.success("Registrado!")

elif menu == "📦 Entrega (Saída)":
    st.header("📤 Entrega de Fardamento")
    col1, col2, col3 = st.columns(3)
    with col1: it = st.selectbox("Item", ITENS)
    with col2: co = st.selectbox("Cor", CORES)
    with col3: ta = st.selectbox("Tamanho", TAMANHOS)
    
    mod = st.selectbox("Modelo", MODELOS)
    est = st.radio("Origem:", ["Novo", "Usado/Reuso"], horizontal=True)

    saldo = consultar_saldo(it, co, ta, mod, est)
    if saldo > 0: st.success(f"✅ Saldo disponível: {saldo}")
    else: st.error("❌ Estoque ZERADO")

    with st.form("confirmar_saida"):
        dest = st.text_input("Recebedor").upper()
        qt_saida = st.number_input("Quantidade", min_value=1, max_value=max(1, int(saldo)))
        if st.form_submit_button("Finalizar Entrega"):
            if saldo >= qt_saida and dest:
                with conectar() as conn:
                    conn.execute("INSERT INTO estoque (item, cor, tamanho, modelo, qtd, estado, pessoa, data_hora, almoxarife) VALUES (?,?,?,?,?,?,?,?,?)", (it, co, ta, mod, -qt_saida, est, dest, datetime.now().strftime("%d/%m/%Y %H:%M"), st.session_state.almoxarife))
                st.success("Saída registrada!"); st.rerun()

elif menu == "📊 Relatório Geral":
    st.header("📊 Movimentação Completa")
    with conectar() as conn:
        df_geral = pd.read_sql_query("SELECT id, item, cor, tamanho, modelo, qtd, estado, data_hora, almoxarife FROM estoque ORDER BY id DESC", conn)
    st.info("Nota: Colunas sensíveis (NF e Recebedor) foram ocultadas nesta visualização pública.")
    st.dataframe(df_geral, use_container_width=True)
