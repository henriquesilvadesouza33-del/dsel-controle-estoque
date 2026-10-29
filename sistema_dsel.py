import base64
import streamlit as st
import psycopg2
import psycopg2.errors
import pandas as pd
from datetime import datetime, date
import os
import io
import unicodedata
import hashlib
import secrets
from docx import Document
from fpdf import FPDF
try:
    from PIL import Image
    PIL_DISPONIVEL = True
except ImportError:
    PIL_DISPONIVEL = False

# --- 1. CONFIGURACOES E ESTILO ---
st.set_page_config(page_title="DSEL - GESTÃO NOVACAP", layout="wide")
NOME_LOGO = "logo.novacap.png"
NOME_LOGO_PDF = "NOVACAP_BRASILIA_LOGO.png"

st.markdown("""
    <style>
    .stMetric { background-color: #1e2630; padding: 15px; border-radius: 10px; border-left: 5px solid #0056b3; }
    .stButton>button { width: 100%; font-weight: bold; border-radius: 8px; height: 3em; }
    .stDataFrame { background-color: #1e2630; border-radius: 10px; }
    .footer-henrique {
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background-color: #0e1117;
        color: #888;
        text-align: center;
        padding: 10px;
        border-top: 1px solid #333;
        font-size: 13px;
        z-index: 999;
    }
    </style>
""", unsafe_allow_html=True)


def rodape_henrique():
    st.markdown("""
        <div class="footer-henrique">
            Este sistema foi desenvolvido por Henrique Silva de Souza
        </div>
    """, unsafe_allow_html=True)


# --- 2. BANCO DE DADOS (PostgreSQL / Supabase) ---
def conectar():
    # Puxa as credenciais ocultas do painel do Streamlit Cloud de forma segura.
    # IMPORTANTE: use as credenciais do BANCO (Session Pooler), nao as API Keys.
    return psycopg2.connect(
        host=st.secrets["SUPABASE_HOST"],
        dbname=st.secrets["SUPABASE_DB"],
        user=st.secrets["SUPABASE_USER"],
        password=st.secrets["SUPABASE_PASSWORD"],
        port=st.secrets["SUPABASE_PORT"],
        sslmode="require",
    )


def executar(sql, params=None):
    """Executa INSERT/UPDATE/DELETE com commit."""
    with conectar() as conn:
        with conn.cursor() as c:
            c.execute(sql, params or ())
        conn.commit()


def buscar_um(sql, params=None):
    """Executa um SELECT e devolve a primeira linha (ou None)."""
    with conectar() as conn:
        with conn.cursor() as c:
            c.execute(sql, params or ())
            return c.fetchone()


# --- SEGURANCA DE SENHAS (LGPD) ---
PREFIXO_HASH = "pbkdf2_sha256"
ITERACOES_HASH = 260000


def hash_senha(senha):
    """Gera um hash seguro (PBKDF2-SHA256) com salt aleatorio."""
    salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", str(senha).encode("utf-8"),
                             salt.encode("utf-8"), ITERACOES_HASH)
    return f"{PREFIXO_HASH}${ITERACOES_HASH}${salt}${dk.hex()}"


def eh_hash(valor):
    return isinstance(valor, str) and valor.startswith(PREFIXO_HASH + "$")


def verificar_senha(senha_digitada, armazenada):
    """Compara a senha digitada com o valor guardado (hash novo ou texto antigo)."""
    if armazenada is None:
        return False
    if eh_hash(armazenada):
        try:
            _, iteracoes, salt, hash_guardado = armazenada.split("$", 3)
            dk = hashlib.pbkdf2_hmac("sha256", str(senha_digitada).encode("utf-8"),
                                     salt.encode("utf-8"), int(iteracoes))
            return secrets.compare_digest(dk.hex(), hash_guardado)
        except Exception:
            return False
    # Compatibilidade: senha antiga guardada em texto puro
    return secrets.compare_digest(str(senha_digitada), str(armazenada))


def inicializar_banco():
    with conectar() as conn:
        with conn.cursor() as c:
            c.execute('''CREATE TABLE IF NOT EXISTS usuarios (
                            id SERIAL PRIMARY KEY,
                            nome TEXT UNIQUE,
                            senha TEXT,
                            status TEXT)''')
            c.execute('''CREATE TABLE IF NOT EXISTS estoque (
                            id SERIAL PRIMARY KEY,
                            item TEXT,
                            cor TEXT,
                            tamanho TEXT,
                            modelo TEXT,
                            qtd INTEGER,
                            estado TEXT,
                            nf TEXT,
                            pessoa TEXT,
                            data_hora TEXT,
                            almoxarife TEXT)''')
            c.execute("INSERT INTO usuarios (nome, senha, status) VALUES ('ADMIN', %s, 'Ativo') ON CONFLICT (nome) DO NOTHING", (hash_senha('123456'),))
            c.execute("UPDATE estoque SET modelo = 'BRIM' WHERE modelo IN ('Calca BRIM', 'brim', 'Brim')")
            c.execute("UPDATE estoque SET modelo = 'Padrão' WHERE modelo IN ('Manga Curta', 'Manga Longa')")
        conn.commit()


inicializar_banco()


# --- 3. FUNCOES DE APOIO ---
def exibir_logo(largura=200):
    if os.path.exists(NOME_LOGO):
        st.image(NOME_LOGO, width=largura)


def consultar_saldo(item, cor, tamanho, modelo, estado):
    res = buscar_um(
        "SELECT SUM(qtd) FROM estoque WHERE item=%s AND cor=%s AND tamanho=%s AND modelo=%s AND estado=%s",
        (item, cor, tamanho, modelo, estado),
    )
    return res[0] if res and res[0] is not None else 0


# --- 4. FUNCOES DE EXPORTACAO ---
def gerar_csv(df):
    return df.to_csv(index=False).encode('utf-8-sig')


def gerar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    return output.getvalue()


def gerar_word(df, titulo="Relatório DSEL NOVACAP"):
    doc = Document()
    doc.add_heading(titulo, 0)
    doc.add_paragraph(f"Relatório gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    table = doc.add_table(rows=1, cols=len(df.columns))
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(df.columns):
        hdr_cells[i].text = str(col)
    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            row_cells[i].text = str(value) if value is not None else ""
    doc_io = io.BytesIO()
    doc.save(doc_io)
    return doc_io.getvalue()


def remover_acentos(texto):
    """Remove acentos para compatibilidade com FPDF (latin-1)."""
    if not isinstance(texto, str):
        texto = str(texto)
    mapeamento = {
        '\u00e7': 'c', '\u00c7': 'C',
        '\u00e3': 'a', '\u00c3': 'A',
        '\u00f5': 'o', '\u00d5': 'O',
        '\u00e1': 'a', '\u00c1': 'A',
        '\u00e9': 'e', '\u00c9': 'E',
        '\u00ed': 'i', '\u00cd': 'I',
        '\u00f3': 'o', '\u00d3': 'O',
        '\u00fa': 'u', '\u00da': 'U',
        '\u00e2': 'a', '\u00c2': 'A',
        '\u00ea': 'e', '\u00ca': 'E',
        '\u00f4': 'o', '\u00d4': 'O',
        '\u00e0': 'a', '\u00c0': 'A',
        '\u00f1': 'n', '\u00d1': 'N',
    }
    for char, subst in mapeamento.items():
        texto = texto.replace(char, subst)
    return "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')


def gerar_pdf(df, titulo="Relatório DSEL NOVACAP"):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, remover_acentos(titulo), ln=1, align='C')
    pdf.set_font("Arial", '', 8)
    pdf.cell(0, 6, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=1, align='C')
    pdf.ln(4)

    larguras_padrao = {
        'id': 12, 'ID': 12,
        'Data/Hora': 28, 'data_hora': 28,
        'Item': 22, 'item': 22,
        'Modelo': 26, 'modelo': 26,
        'Cor': 18, 'cor': 18,
        'Tam.': 14, 'tamanho': 14,
        'Qtd': 12, 'qtd': 12,
        'Qtd. Retirada': 16,
        'Estado': 18, 'estado': 18,
        'NF': 20, 'nf': 20,
        'Origem/Destino': 48, 'pessoa': 48,
        'Reeducando': 48,
        'Almoxarife': 30, 'almoxarife': 30,
        'Saldo': 14, 'Saldo Atual': 16
    }

    n_cols = len(df.columns)
    largura_total = 277

    col_widths = []
    for col in df.columns:
        col_str = str(col)
        if col_str in larguras_padrao:
            col_widths.append(larguras_padrao[col_str])
        else:
            encontrado = False
            for key, val in larguras_padrao.items():
                if key.lower() in col_str.lower():
                    col_widths.append(val)
                    encontrado = True
                    break
            if not encontrado:
                col_widths.append(max(15, largura_total // n_cols))

    soma = sum(col_widths)
    if soma > largura_total:
        fator = largura_total / soma
        col_widths = [w * fator for w in col_widths]

    pdf.set_font("Arial", 'B', 7)
    for i, col in enumerate(df.columns):
        nome_col = remover_acentos(str(col))
        max_chars = int(col_widths[i] / 1.8)
        pdf.cell(col_widths[i], 6, nome_col[:max_chars], border=1, align='C')
    pdf.ln()

    pdf.set_font("Arial", '', 7)
    for index, row in df.iterrows():
        for i, val in enumerate(row):
            texto_celula = remover_acentos(val) if val is not None else ""
            max_chars = int(col_widths[i] / 1.6)
            texto_truncado = texto_celula[:max_chars] if len(texto_celula) > max_chars else texto_celula
            pdf.cell(col_widths[i], 5, texto_truncado, border=1, align='C')
        pdf.ln()
    return bytes(pdf.output())


def botoes_exportacao(df, nome_base="relatorio", titulo_pdf="Relatório DSEL NOVACAP"):
    col1, col2, col3, col4 = st.columns(4)
    data_str = datetime.now().strftime('%Y%m%d_%H%M')
    with col1:
        st.download_button("\U0001F4CA Excel", gerar_excel(df), f"{nome_base}_{data_str}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col2:
        st.download_button("\U0001F4DD Word", gerar_word(df, titulo_pdf), f"{nome_base}_{data_str}.docx",
                           "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with col3:
        st.download_button("\U0001F4C4 PDF", gerar_pdf(df, titulo_pdf), f"{nome_base}_{data_str}.pdf",
                           "application/pdf")
    with col4:
        st.download_button("\U0001F4CB CSV", gerar_csv(df), f"{nome_base}_{data_str}.csv",
                           "text/csv")


# --- COMPROVANTE DE ENTREGA (PDF INDIVIDUAL) ---
def gerar_comprovante_entrega(nome_recebedor, df_itens, almoxarife_nome):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    logo_pdf = NOME_LOGO_PDF if os.path.exists(NOME_LOGO_PDF) else NOME_LOGO
    if os.path.exists(logo_pdf) and PIL_DISPONIVEL:
        try:
            img = Image.open(logo_pdf)
            if img.mode in ('RGBA', 'LA', 'P'):
                img = img.convert('RGB')
            temp_logo = "_temp_logo_novacap.png"
            img.save(temp_logo, format='PNG')
            pdf.image(temp_logo, x=10, y=8, w=30)
            if os.path.exists(temp_logo):
                os.remove(temp_logo)
        except Exception:
            pass

    def t(txt):
        return str(txt).encode('latin1', 'replace').decode('latin1')

    pdf.set_y(12)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 8, t("NOVACAP - DSEL"), ln=1, align='C')
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 6, t("DIVISÃO DE SEGURANÇA, LIMPEZA E CONSERVAÇÃO"), ln=1, align='C')
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 6, t("COMPROVANTE DE ENTREGA DE MATERIAL"), ln=1, align='C')
    pdf.ln(2)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(6)

    pdf.set_font("Arial", 'B', 10)
    pdf.cell(60, 8, t("Recebedor:"), ln=0)
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 8, t(nome_recebedor), ln=1)

    pdf.set_font("Arial", 'B', 10)
    pdf.cell(60, 8, t("Data da Entrega:"), ln=0)
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 8, date.today().strftime("%d/%m/%Y"), ln=1)

    pdf.set_font("Arial", 'B', 10)
    pdf.cell(60, 8, t("Responsável pelo Lançamento:"), ln=0)
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 8, t(almoxarife_nome), ln=1)
    pdf.ln(5)

    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(220, 220, 220)
    pdf.cell(60, 8, t("Item"), border=1, align='C', fill=True)
    pdf.cell(35, 8, t("Modelo"), border=1, align='C', fill=True)
    pdf.cell(30, 8, t("Cor"), border=1, align='C', fill=True)
    pdf.cell(25, 8, t("Tam."), border=1, align='C', fill=True)
    pdf.cell(20, 8, t("Qtd"), border=1, align='C', fill=True)
    pdf.cell(20, 8, t("Estado"), border=1, align='C', fill=True)
    pdf.ln()

    pdf.set_font("Arial", '', 9)
    for _, row in df_itens.iterrows():
        item_texto = str(row.get('Item', ''))
        estado_texto = str(row.get('Estado', ''))
        pdf.cell(60, 7, t(item_texto[:28]), border=1, align='L')
        pdf.cell(35, 7, t(str(row.get('Modelo', ''))[:16]), border=1, align='L')
        pdf.cell(30, 7, t(str(row.get('Cor', ''))[:14]), border=1, align='L')
        pdf.cell(25, 7, t(str(row.get('Tam.', ''))[:10]), border=1, align='C')
        pdf.cell(20, 7, str(abs(int(row.get('Qtd', 0)))), border=1, align='C')
        pdf.cell(20, 7, t(estado_texto[:10]), border=1, align='C')
        pdf.ln()

    pdf.ln(10)
    pdf.set_font("Arial", '', 9)
    texto_declaracao = ("Declaro que recebi o material descrito acima em perfeitas condições e me "
                        "comprometo a utiliza-lo exclusivamente para as atividades relacionadas as "
                        "minhas funções na NOVACAP.")
    pdf.multi_cell(0, 5, t(texto_declaracao))
    pdf.ln(15)

    y_assinatura = pdf.get_y()
    pdf.line(25, y_assinatura, 95, y_assinatura)
    pdf.line(115, y_assinatura, 185, y_assinatura)
    pdf.ln(3)
    pdf.set_font("Arial", '', 9)
    pdf.cell(95, 5, t("Assinatura do Recebedor"), align='C')
    pdf.cell(10, 5, "", align='C')
    pdf.cell(95, 5, t("Assinatura do Almoxarife"), align='C')
    pdf.ln(1)

    pdf.set_y(-20)
    pdf.set_font("Arial", 'I', 8)
    texto_rodape = f"Sistema desenvolvido por Henrique Silva de Souza - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    pdf.cell(0, 5, t(texto_rodape), ln=1, align='C')

    return bytes(pdf.output())


def mostrar_pdf_inline(comprovante_pdf, altura=700):
    if isinstance(comprovante_pdf, str):
        with open(comprovante_pdf, "rb") as f:
            pdf_bytes = f.read()
    else:
        pdf_bytes = bytes(comprovante_pdf)
    base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
    pdf_display = (f'<iframe src="data:application/pdf;base64,{base64_pdf}" '
                   f'width="100%" height="{altura}" type="application/pdf"></iframe>')
    st.markdown(pdf_display, unsafe_allow_html=True)


# --- 5. CONTROLE DE ACESSO ---
if 'logado' not in st.session_state:
    st.session_state.logado = False

if not st.session_state.logado:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        exibir_logo(250)
        st.title("\U0001F510 Acesso DSEL")
        u = st.text_input("Usuário").strip().upper()
        s = st.text_input("Senha", type="password")
        if st.button("Entrar"):
            res = buscar_um(
                "SELECT nome, senha FROM usuarios WHERE UPPER(nome)=%s AND status='Ativo'",
                (u,),
            )
            if res and verificar_senha(s, res[1]):
                # Atualiza automaticamente senhas antigas (texto puro) para hash
                if not eh_hash(res[1]):
                    executar("UPDATE usuarios SET senha=%s WHERE nome=%s", (hash_senha(s), res[0]))
                st.session_state.logado, st.session_state.almoxarife = True, res[0]
                # Forca troca se ainda estiver usando a senha padrao
                st.session_state.precisa_trocar_senha = (s == "123456")
                st.rerun()
            else:
                st.error("Acesso negado.")
    rodape_henrique()
else:
    CORES = ["Cinza", "Amarelo", "Azul", "Branco", "Verde", "Preto", "Marrom", "Laranja", "Vermelho"]
    ITENS = ["Camiseta", "Calça", "Bota", "Luvas", "Chapéu", "Boné", "Cinto", "Meias"]
    TAMANHOS = ["Único", "P", "M", "G", "GG", "XG", "XGG"] + [str(i).zfill(2) for i in range(34, 57)]
    MODELOS = ["BRIM", "JEANS", "Padrão", "Botina", "Bontina Galocha", "luva de couro", "luva vaqueta",
               "luvas de pvc cano longo", "Cano Alto", "Refletivo", "Elastano", "Tactel", "Camuflado"]

    with st.sidebar:
        exibir_logo(150)
        st.subheader(f"\U0001F464 {st.session_state.almoxarife}")
        opcoes = [
            "\U0001F3E0 Painel Principal",
            "\U0001F4E6 Saldo Geral",
            "\U0001F50D Consulta Individual",
            "\U0001F4E5 Entrada (Fábrica)",
            "\U0001F4E6 Entrega (Saída)",
            "\U0001F504 Retorno/Devolução",
            "\U0001F4D1 Extrato por Item",
            "\U0001F4CB Extrato Geral",
            "\U0001F4CA Relatório Geral",
            "\U0001F4C5 Lançamentos Diários",
            "\U0001F511 Minha Senha"
        ]
        if st.session_state.almoxarife == "ADMIN":
            opcoes.append("\u2699\uFE0F Gestão de Lançamentos")
            opcoes.append("\U0001F527 Ajuste de Inventário")
            opcoes.append("\U0001F195 Novo Cadastro")
            opcoes.append("\u270F\uFE0F Editar Usuários")
        menu = st.radio("Navegação", opcoes)
        if st.button("Sair"):
            st.session_state.logado = False
            st.rerun()

    # --- Troca obrigatoria da senha padrao (LGPD) ---
    if st.session_state.get("precisa_trocar_senha"):
        st.warning("\u26A0\uFE0F Por seguranca (LGPD), defina uma nova senha antes de continuar. A senha padrao nao pode ser mantida.")
        with st.form("form_troca_obrigatoria"):
            nova1 = st.text_input("Nova senha", type="password")
            nova2 = st.text_input("Confirme a nova senha", type="password")
            if st.form_submit_button("Salvar nova senha"):
                if not nova1 or len(nova1) < 6:
                    st.error("A senha precisa ter pelo menos 6 caracteres.")
                elif nova1 != nova2:
                    st.error("As senhas nao conferem.")
                elif nova1 == "123456":
                    st.error("Escolha uma senha diferente da padrao.")
                else:
                    executar("UPDATE usuarios SET senha=%s WHERE nome=%s", (hash_senha(nova1), st.session_state.almoxarife))
                    st.session_state.precisa_trocar_senha = False
                    st.success("Senha atualizada! O sistema sera liberado.")
                    st.rerun()
        rodape_henrique()
        st.stop()

    # --- 6. MODULOS ---
    if "Painel Principal" in menu:
        st.header("\U0001F3E0 Detalhamento de Itens em Estoque")
        item_sel = st.selectbox("Selecione um item:", ["-- Selecione --"] + ITENS)
        if item_sel != "-- Selecione --":
            with conectar() as conn:
                df_det = pd.read_sql_query("SELECT modelo, cor, tamanho, estado, SUM(qtd) as \"Saldo\" FROM estoque WHERE item=%s GROUP BY modelo, cor, tamanho, estado HAVING SUM(qtd) != 0", conn, params=(item_sel,))
            if not df_det.empty:
                st.dataframe(df_det, use_container_width=True, hide_index=True)
                botoes_exportacao(df_det, f"saldo_{item_sel}", f"Saldo Detalhado - {item_sel}")
            else:
                st.info("Nenhum saldo encontrado para este item.")
        rodape_henrique()

    elif "Saldo Geral" in menu:
        st.header("\U0001F4E6 Saldo Geral Disponível")
        with conectar() as conn:
            df_saldo = pd.read_sql_query("SELECT item as \"Item\", modelo as \"Modelo\", cor as \"Cor\", tamanho as \"Tam.\", estado as \"Estado\", SUM(qtd) as \"Saldo Atual\" FROM estoque GROUP BY item, modelo, cor, tamanho, estado HAVING SUM(qtd) != 0 ORDER BY item ASC", conn)
        if not df_saldo.empty:
            st.dataframe(df_saldo, use_container_width=True, hide_index=True)
            st.info(f"\U0001F4A1 Total de peças físicas atualmente no estoque: **{df_saldo['Saldo Atual'].sum()}**")
            botoes_exportacao(df_saldo, "saldo_geral", "Saldo Geral de Estoque")
        else:
            st.info("Nenhum item com saldo positivo no estoque.")
        rodape_henrique()

    elif "Consulta Individual" in menu:
        st.header("\U0001F50D Consulta por Reeducando")
        busca = st.text_input("Digite o Nome, CPF ou Matrícula:").strip().upper()
        if busca:
            with conectar() as conn:
                query = "SELECT data_hora as \"Data/Hora\", item as \"Item\", modelo as \"Modelo\", cor as \"Cor\", tamanho as \"Tam.\", ABS(qtd) as \"Qtd. Retirada\", estado as \"Estado\", pessoa as \"Reeducando\", almoxarife as \"Almoxarife\" FROM estoque WHERE UPPER(pessoa) LIKE %s AND qtd < 0 ORDER BY id DESC"
                df_reeducando = pd.read_sql_query(query, conn, params=(f"%{busca}%",))
            if not df_reeducando.empty:
                st.dataframe(df_reeducando, use_container_width=True, hide_index=True)
                st.info(f"\U0001F4A1 Total de itens retirados: **{df_reeducando['Qtd. Retirada'].sum()}**")
                botoes_exportacao(df_reeducando, f"consulta_{busca}", "Consulta Individual de Reeducando")
            else:
                st.warning("Nenhum registro encontrado.")
        rodape_henrique()

    elif "Entrada (Fábrica)" in menu:
        st.header("\U0001F4E5 Entrada da Fábrica")
        with st.form("ent"):
            nf = st.text_input("Nota Fiscal")
            it, co, ta, mod = st.selectbox("Item", ITENS), st.selectbox("Cor", CORES), st.selectbox("Tamanho", TAMANHOS), st.selectbox("Modelo", MODELOS)
            qt = st.number_input("Quantidade", min_value=1)
            if st.form_submit_button("Salvar Entrada"):
                executar("INSERT INTO estoque (item, cor, tamanho, modelo, qtd, estado, nf, pessoa, data_hora, almoxarife) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                         (it, co, ta, mod, qt, "Novo", nf, "Fabrica", datetime.now().strftime("%d/%m/%Y %H:%M"), st.session_state.almoxarife))
                st.success("Registrado!")
                st.rerun()
        rodape_henrique()

    elif "Entrega (Saída)" in menu:
        st.header("\U0001F4E4 Entrega de Fardamento")
        col1, col2, col3 = st.columns(3)
        with col1: it = st.selectbox("Item", ITENS, index=ITENS.index("Calça") if "Calça" in ITENS else 0)
        with col2: co = st.selectbox("Cor", CORES)
        with col3: ta = st.selectbox("Tamanho", TAMANHOS)
        col4, col5 = st.columns(2)
        with col4: mod = st.selectbox("Modelo", MODELOS)
        with col5: est = st.radio("Origem:", ["Novo", "Usado/Reuso"], horizontal=True)

        saldo = consultar_saldo(it, co, ta, mod, est)
        if saldo > 0:
            st.success(f"\u2705 Saldo disponível: {saldo}")
        else:
            st.error(f"\u274C Estoque ZERADO para {it} {mod}")

        with st.form("confirmar_saida"):
            dest = st.text_input("Recebedor").upper()
            qt_saida = st.number_input("Quantidade", min_value=1, max_value=max(1, int(saldo)))
            if st.form_submit_button("Finalizar Entrega"):
                saldo_atualizado = consultar_saldo(it, co, ta, mod, est)
                if saldo_atualizado <= 0 or qt_saida > saldo_atualizado:
                    st.error("Não há saldo suficiente disponível no momento do envio!")
                elif not dest:
                    st.warning("Preencha o recebedor.")
                else:
                    executar("INSERT INTO estoque (item, cor, tamanho, modelo, qtd, estado, pessoa, data_hora, almoxarife) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                             (it, co, ta, mod, -int(qt_saida), est, dest, datetime.now().strftime("%d/%m/%Y %H:%M"), st.session_state.almoxarife))
                    st.success("Saída registrada com sucesso!")
                    st.rerun()
        rodape_henrique()

    elif "Retorno/Devolução" in menu:
        st.header("\U0001F504 Retorno / Devolução de Item")
        with st.form("devolucao"):
            it, co, ta, mod = st.selectbox("Item", ITENS), st.selectbox("Cor", CORES), st.selectbox("Tamanho", TAMANHOS), st.selectbox("Modelo", MODELOS)
            est = st.radio("Estado de retorno:", ["Usado/Reuso", "Novo"], horizontal=True)
            qt_ret = st.number_input("Quantidade Retornando", min_value=1)
            origem = st.text_input("Quem está devolvendo").upper()
            if st.form_submit_button("Registrar Retorno"):
                if not origem:
                    st.warning("Informe o reeducando.")
                else:
                    executar("INSERT INTO estoque (item, cor, tamanho, modelo, qtd, estado, pessoa, data_hora, almoxarife) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                             (it, co, ta, mod, int(qt_ret), est, f"Devolução: {origem}", datetime.now().strftime("%d/%m/%Y %H:%M"), st.session_state.almoxarife))
                    st.success("Retorno registrado!")
                    st.rerun()
        rodape_henrique()

    elif "Extrato por Item" in menu:
        st.header("\U0001F4D1 Extrato de Movimentação por Item")
        col1, col2, col3 = st.columns(3)
        with col1: it = st.selectbox("Item", ITENS)
        with col2: co = st.selectbox("Cor", CORES)
        with col3: ta = st.selectbox("Tamanho", TAMANHOS)
        col4, col5 = st.columns(2)
        with col4: mod = st.selectbox("Modelo", MODELOS)
        with col5: est = st.radio("Estado:", ["Novo", "Usado/Reuso"], horizontal=True)

        with conectar() as conn:
            df_extrato = pd.read_sql_query("SELECT data_hora as \"Data/Hora\", nf as \"NF\", pessoa as \"Origem/Destino\", qtd as \"Qtd\" FROM estoque WHERE item=%s AND cor=%s AND tamanho=%s AND modelo=%s AND estado=%s ORDER BY id ASC", conn, params=(it, co, ta, mod, est))

        if not df_extrato.empty:
            df_extrato['Saldo Acumulado'] = df_extrato['Qtd'].cumsum()
            st.dataframe(df_extrato, use_container_width=True, hide_index=True)
            st.metric(label="Saldo Final Disponível", value=f"{df_extrato['Saldo Acumulado'].iloc[-1]} unidades")
            botoes_exportacao(df_extrato, f"extrato_{it}_{mod}", f"Extrato por Item - {it} {mod}")
        else:
            st.info("Nenhuma movimentação encontrada.")
        rodape_henrique()

    elif "Extrato Geral" in menu:
        st.header("\U0001F4CB Extrato Geral de Movimentações")
        st.write("Histórico consolidado de movimentações para visualização e impressão.")
        with conectar() as conn:
            df_extrato_geral = pd.read_sql_query("SELECT data_hora as \"Data/Hora\", item as \"Item\", modelo as \"Modelo\", cor as \"Cor\", tamanho as \"Tam.\", qtd as \"Qtd\", estado as \"Estado\", nf as \"NF\", pessoa as \"Origem/Destino\", almoxarife as \"Almoxarife\" FROM estoque ORDER BY id DESC", conn)

        if not df_extrato_geral.empty:
            botoes_exportacao(df_extrato_geral, "extrato_geral", "Extrato Geral de Movimentações")
            st.write("---")
            ver_preview = st.checkbox("\U0001F441\uFE0F Visualizar como ficara impresso")
            if ver_preview:
                st.subheader("\U0001F4C4 Pré-visualização da Impressão")
                st.markdown("<div style='background-color: #ffffff; color: #111111; padding: 25px; border-radius: 5px; border: 2px solid #333; font-family: Courier New, monospace;'><h3 style='text-align: center; margin-bottom: 0;'>NOVACAP - DIVISÃO DE SEGURANÇA, LIMPEZA E CONSERVAÇÃO (DSEL)</h3><p style='text-align: center; margin-top: 5px;'><b>Relatório consolidado de Extrato Geral</b></p><p style='text-align: right; font-size: 12px;'>Data de emissão: " + datetime.now().strftime('%d/%m/%Y %H:%M') + "</p><hr style='border-top: 1px dashed black; margin-bottom: 15px;'></div>", unsafe_allow_html=True)
                st.dataframe(df_extrato_geral, use_container_width=True, hide_index=True)
            st.write("### \U0001F4CB Tabela de Dados no Sistema")
            st.dataframe(df_extrato_geral, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma movimentacao registrada no sistema.")
        rodape_henrique()

    elif "Relatório Geral" in menu:
        st.header("\U0001F4CA Movimentação Completa (Tabela do Banco)")
        with conectar() as conn:
            df_geral = pd.read_sql_query("SELECT * FROM estoque ORDER BY id DESC", conn)
        if not df_geral.empty:
            st.dataframe(df_geral, use_container_width=True, hide_index=True)
            botoes_exportacao(df_geral, "relatorio_geral", "Relatório Geral Completo")
        else:
            st.info("Nenhuma movimentacao registrada.")
        rodape_henrique()

    elif "Lançamentos Diários" in menu:
        st.header("\U0001F4C5 Lançamentos do Dia")
        data_calendario = st.date_input("Selecione a data para filtrar:", date.today(), format="DD/MM/YYYY")
        data_sel_str = data_calendario.strftime("%d/%m/%Y")
        st.subheader(f"Data: {data_sel_str}")

        with conectar() as conn:
            df_diario = pd.read_sql_query(
                "SELECT data_hora as \"Data/Hora\", item as \"Item\", modelo as \"Modelo\", cor as \"Cor\", tamanho as \"Tam.\", qtd as \"Qtd\", estado as \"Estado\", nf as \"NF\", pessoa as \"Origem/Destino\", almoxarife as \"Almoxarife\" FROM estoque WHERE data_hora LIKE %s ORDER BY id DESC",
                conn,
                params=(f"{data_sel_str}%",)
            )

        if not df_diario.empty:
            st.success(f"\u2705 {len(df_diario)} lançamento(s) encontrado(s) para {data_sel_str}.")
            st.dataframe(df_diario, use_container_width=True, hide_index=True)

            entradas = df_diario[df_diario['Qtd'] > 0]['Qtd'].sum()
            saidas = abs(df_diario[df_diario['Qtd'] < 0]['Qtd'].sum())
            col1, col2 = st.columns(2)
            col1.metric("\U0001F4E5 Total de Entradas", f"{entradas} un")
            col2.metric("\U0001F4E4 Total de Saídas", f"{saidas} un")

            botoes_exportacao(df_diario, f"lancamentos_diarios_{data_sel_str.replace('/','')}", "Lançamentos Diários - DSEL")

            st.divider()
            st.subheader("\U0001F5A8\uFE0F Comprovante de Entrega por Pessoa")
            st.write(f"Selecione uma pessoa para gerar um comprovante de entrega individualizado (PDF) com os itens que ela recebeu em {data_sel_str}.")

            saidas_hoje = df_diario[df_diario['Qtd'] < 0]
            saidas_hoje = saidas_hoje[
                ~saidas_hoje['Origem/Destino'].astype(str).str.upper().apply(remover_acentos).isin(['FABRICA', 'AJUSTE']) &
                ~saidas_hoje['Origem/Destino'].astype(str).str.upper().apply(remover_acentos).str.startswith('DEVOLUCAO')
            ]

            if not saidas_hoje.empty:
                pessoas_unicas = sorted(saidas_hoje['Origem/Destino'].dropna().unique())
                pessoa_sel = st.selectbox("Selecione o recebedor:", ["-- Selecione --"] + list(pessoas_unicas))

                if pessoa_sel != "-- Selecione --":
                    df_pessoa = saidas_hoje[saidas_hoje['Origem/Destino'] == pessoa_sel][
                        ['Item', 'Modelo', 'Cor', 'Tam.', 'Qtd', 'Estado']
                    ].copy()
                    df_pessoa['Qtd'] = df_pessoa['Qtd'].abs()

                    st.write(f"**Itens entregues em {data_sel_str} para {pessoa_sel}:**")
                    st.dataframe(df_pessoa, use_container_width=True, hide_index=True)

                    comprovante_pdf = gerar_comprovante_entrega(pessoa_sel, df_pessoa, st.session_state.almoxarife)

                    st.write("---")
                    st.subheader("\U0001F441\uFE0F Pré-visualização do Comprovante")
                    mostrar_pdf_inline(comprovante_pdf, altura=700)

                    st.download_button(
                        label="\U0001F4C4 Baixar Comprovante de Entrega (PDF)",
                        data=comprovante_pdf,
                        file_name=f"comprovante_entrega_{pessoa_sel.replace(' ','_')}_{data_sel_str.replace('/','')}.pdf",
                        mime="application/pdf"
                    )

                    st.info("\U0001F4A1 O PDF gerado inclui: logo NOVACAP, cabeçalho DSEL, lista de itens, data, nome do almoxarife, declaração de recebimento, campos para assinatura e rodapé com crédito do desenvolvedor.")
            else:
                st.info(f"Nenhuma entrega (saída) registrada em {data_sel_str} para geração de comprovantes.")
        else:
            st.info(f"Nenhum lançamento registrado em {data_sel_str}.")
        rodape_henrique()

    elif "Minha Senha" in menu:
        st.header("\U0001F511 Trocar Minha Senha")
        with st.form("form_minha_senha"):
            atual = st.text_input("Senha atual", type="password")
            nova1 = st.text_input("Nova senha", type="password")
            nova2 = st.text_input("Confirme a nova senha", type="password")
            if st.form_submit_button("Alterar senha"):
                reg = buscar_um("SELECT senha FROM usuarios WHERE nome=%s", (st.session_state.almoxarife,))
                if not reg or not verificar_senha(atual, reg[0]):
                    st.error("Senha atual incorreta.")
                elif len(nova1) < 6:
                    st.error("A nova senha precisa ter pelo menos 6 caracteres.")
                elif nova1 != nova2:
                    st.error("As senhas nao conferem.")
                else:
                    executar("UPDATE usuarios SET senha=%s WHERE nome=%s", (hash_senha(nova1), st.session_state.almoxarife))
                    st.success("Senha alterada com sucesso!")
        rodape_henrique()

    elif "Gestão de Lançamentos" in menu and st.session_state.almoxarife == "ADMIN":
        st.header("\u2699\uFE0F Gestão de Lançamentos (Modo Supervisor)")
        st.warning("\u26A0\uFE0F Atenção: A exclusão ou edição de registros altera o saldo do estoque permanentemente.")
        with conectar() as conn:
            df_admin = pd.read_sql_query("SELECT * FROM estoque ORDER BY id DESC", conn)

        if not df_admin.empty:
            st.subheader("\U0001F50D 1. Encontrar Lançamento")
            termo_busca = st.text_input("Digite o nome do Recebedor/Reeducando, ID, NF ou Item para filtrar:").strip().upper()

            if termo_busca:
                df_filtrado = df_admin[
                    df_admin['pessoa'].astype(str).str.upper().str.contains(termo_busca, na=False) |
                    df_admin['item'].astype(str).str.upper().str.contains(termo_busca, na=False) |
                    df_admin['nf'].astype(str).str.upper().str.contains(termo_busca, na=False) |
                    df_admin['id'].astype(str).str.contains(termo_busca, na=False)
                ]
            else:
                df_filtrado = df_admin

            if not df_filtrado.empty:
                lista_opcoes = {
                    f"ID: {row['id']} | Data: {row['data_hora']} | {row['item']} ({row['modelo']}) | Qtd: {row['qtd']} | Destino: {row['pessoa']} | Estado: {row['estado']}": row['id']
                    for idx, row in df_filtrado.iterrows()
                }
                selecionado_rotulo = st.selectbox("Selecione o registro exato para alterar:", ["-- Selecione --"] + list(lista_opcoes.keys()))

                if selecionado_rotulo != "-- Selecione --":
                    id_registro_banco = lista_opcoes[selecionado_rotulo]
                    reg_atual = df_admin[df_admin['id'] == id_registro_banco].iloc[0]
                    st.divider()
                    st.subheader("\U0001F4DD Formulário de Alteração - Registro Interno")

                    col_form1, col_form2, col_form3 = st.columns(3)
                    with col_form1:
                        novo_item = st.selectbox("Corrigir Item:", ITENS, index=ITENS.index(reg_atual['item']) if reg_atual['item'] in ITENS else 0)
                        nova_cor = st.selectbox("Corrigir Cor:", CORES, index=CORES.index(reg_atual['cor']) if reg_atual['cor'] in CORES else 0)
                    with col_form2:
                        novo_modelo = st.selectbox("Corrigir Modelo:", MODELOS, index=MODELOS.index(reg_atual['modelo']) if reg_atual['modelo'] in MODELOS else 0)
                        novo_tamanho = st.selectbox("Corrigir Tam:", TAMANHOS, index=TAMANHOS.index(reg_atual['tamanho']) if reg_atual['tamanho'] in TAMANHOS else 0)
                    with col_form3:
                        nova_qtd = st.number_input("Corrigir Qtd (Valores negativos indicam Saídas):", value=int(reg_atual['qtd']))
                        novo_destino = st.text_input("Corrigir Responsável/Destino:", value=str(reg_atual['pessoa'])).upper()

                    st.write("---")
                    st.subheader("\U0001F504 Alterar Estado do Item")
                    estado_atual = reg_atual['estado'] if reg_atual['estado'] else "Novo"
                    novo_estado = st.radio(
                        "Selecione o novo estado:",
                        ["Novo", "Usado/Reuso"],
                        horizontal=True,
                        index=0 if estado_atual == "Novo" else 1
                    )
                    if estado_atual != novo_estado:
                        st.info(f"O estado sera alterado de **{estado_atual}** para **{novo_estado}**")

                    st.write("---")
                    col_btn1, col_btn2, col_btn3 = st.columns(3)

                    with col_btn1:
                        if st.button("\U0001F4BE Salvar Alterações", type="primary"):
                            executar("UPDATE estoque SET item=%s, modelo=%s, cor=%s, tamanho=%s, qtd=%s, pessoa=%s, estado=%s WHERE id=%s",
                                     (novo_item, novo_modelo, nova_cor, novo_tamanho, int(nova_qtd), novo_destino, novo_estado, int(id_registro_banco)))
                            st.success("Lançamento atualizado com sucesso!")
                            st.rerun()

                    with col_btn2:
                        if st.button("\U0001F6A8 Confirmar Exclusão Definitiva", type="secondary"):
                            executar("DELETE FROM estoque WHERE id = %s", (int(id_registro_banco),))
                            st.warning("Lançamento apagado permanentemente.")
                            st.rerun()

                    with col_btn3:
                        if st.button("\u274C Cancelar Operação"):
                            st.info("Ação cancelada.")
                            st.rerun()
            else:
                st.error("Nenhum lançamento corresponde aos termos pesquisados.")
            st.write("### \U0001F4CB Histórico Completo de Auditoria")
            st.dataframe(df_admin, use_container_width=True, hide_index=True)
            botoes_exportacao(df_admin, "gestao_lancamentos", "Gestão de Lançamentos - Auditoria")
        else:
            st.info("Nenhuma movimentacao registrada no sistema.")
        rodape_henrique()

    elif "Ajuste de Inventário" in menu and st.session_state.almoxarife == "ADMIN":
        st.header("\U0001F527 Ajuste de Balanço de Inventário")
        st.write("Utilize esta aba para forçar o saldo do sistema a bater exatamente com o estoque físico contado.")

        col_adj1, col_adj2, col_adj3 = st.columns(3)
        with col_adj1: item_adj = st.selectbox("Item para Ajustar", ITENS)
        with col_adj2: cor_adj = st.selectbox("Cor para Ajustar", CORES)
        with col_adj3: tam_adj = st.selectbox("Tam para Ajustar", TAMANHOS)

        col_adj4, col_adj5 = st.columns(2)
        with col_adj4: mod_adj = st.selectbox("Modelo para Ajustar", MODELOS)
        with col_adj5: est_adj = st.radio("Estado do Item:", ["Novo", "Usado/Reuso"], horizontal=True)

        saldo_sistema = consultar_saldo(item_adj, cor_adj, tam_adj, mod_adj, est_adj)
        st.metric(label="Saldo Atual no Sistema", value=f"{saldo_sistema} unidades")

        with st.form("form_ajuste"):
            qtd_real_fisica = st.number_input("Quantidade Real Encontrada no Armário (Físico):", min_value=0, value=int(saldo_sistema))
            motivo_ajuste = st.text_input("Motivo do Ajuste (Ex: Inventário Geral 2026, Correção de quebra)").upper()

            if st.form_submit_button("\U0001F528 Aplicar Ajuste de Estoque"):
                diferenca = qtd_real_fisica - saldo_sistema
                if diferenca == 0:
                    st.info("O valor digitado é igual ao do sistema. Nenhum ajuste necessário.")
                elif not motivo_ajuste:
                    st.warning("Descreva o motivo do ajuste para fins de auditoria.")
                else:
                    executar("INSERT INTO estoque (item, cor, tamanho, modelo, qtd, estado, nf, pessoa, data_hora, almoxarife) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                             (item_adj, cor_adj, tam_adj, mod_adj, int(diferenca), est_adj, "AJUSTE", f"INV: {motivo_ajuste}",
                              datetime.now().strftime("%d/%m/%Y %H:%M"), st.session_state.almoxarife))
                    st.success(f"Estoque ajustado! Lançamento de balanceamento de {diferenca} unidades registrado.")
                    st.rerun()
        rodape_henrique()

    elif "Novo Cadastro" in menu and st.session_state.almoxarife == "ADMIN":
        st.header("\U0001F195 Novo Cadastro de Usuários do Sistema")
        with st.form("form_novo_usuario"):
            novo_nome = st.text_input("Nome do Usuário/Almoxarife").strip().upper()
            nova_senha = st.text_input("Senha de Acesso", type="password")
            novo_status = st.selectbox("Status Inicial", ["Ativo", "Inativo"])

            if st.form_submit_button("Salvar Usuário"):
                if not novo_nome or not nova_senha:
                    st.warning("Preencha todos os campos obrigatórios.")
                else:
                    try:
                        executar("INSERT INTO usuarios (nome, senha, status) VALUES (%s, %s, %s)", (novo_nome, hash_senha(nova_senha), novo_status))
                        st.success(f"Usuário {novo_nome} cadastrado com sucesso!")
                    except psycopg2.errors.UniqueViolation:
                        st.error("Erro: Já existe um usuário cadastrado com esse nome.")
                    except Exception as e:
                        st.error(f"Erro ao cadastrar usuário: {e}")
        rodape_henrique()

    elif "Editar Usuários" in menu and st.session_state.almoxarife == "ADMIN":
        st.header("\u270F\uFE0F Gerenciamento e Edição de Usuários")
        with conectar() as conn:
            df_usuarios = pd.read_sql_query("SELECT id, nome, status FROM usuarios ORDER BY nome ASC", conn)

        if not df_usuarios.empty:
            st.subheader("Lista de Usuários Cadastrados")
            st.dataframe(df_usuarios, use_container_width=True, hide_index=True)
            botoes_exportacao(df_usuarios, "usuarios_cadastrados", "Usuários do Sistema")

            st.write("---")
            st.subheader("Alterar Dados de um Usuário")
            usuario_selecionado = st.selectbox("Selecione o usuário para modificar:", df_usuarios["nome"].tolist())

            dados_user = buscar_um("SELECT senha, status FROM usuarios WHERE nome = %s", (usuario_selecionado,))

            if dados_user:
                with st.form("form_edicao_usuario"):
                    ed_senha = st.text_input("Redefinir Senha (deixe em branco para manter a atual)", type="password")
                    ed_status = st.selectbox("Alterar Status", ["Ativo", "Inativo"], index=0 if dados_user[1] == "Ativo" else 1)

                    if st.form_submit_button("Confirmar Alterações"):
                        if ed_senha:
                            if len(ed_senha) < 6:
                                st.error("A nova senha precisa ter pelo menos 6 caracteres.")
                                st.stop()
                            executar("UPDATE usuarios SET senha = %s, status = %s WHERE nome = %s", (hash_senha(ed_senha), ed_status, usuario_selecionado))
                        else:
                            executar("UPDATE usuarios SET status = %s WHERE nome = %s", (ed_status, usuario_selecionado))
                        st.success(f"Dados do usuário {usuario_selecionado} atualizados!")
                        st.rerun()
        else:
            st.info("Nenhum usuario cadastrado no sistema.")
        rodape_henrique()
