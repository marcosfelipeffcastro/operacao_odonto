import streamlit as st
st.set_page_config(page_title="Hapvida + Odonto", layout="wide")
import pandas as pd
import os
import unicodedata
from PIL import Image
import re
import difflib
import base64
from io import BytesIO
from openai import OpenAI

st.set_page_config(
    page_title="Hapvida + Odonto", layout="wide", initial_sidebar_state="collapsed"
)


def to_excel_bytes(dfs):
    """Convert one DataFrame or a dict of {sheet_name: DataFrame} to xlsx bytes."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if isinstance(dfs, pd.DataFrame):
            dfs.to_excel(writer, index=False, sheet_name="Dados")
        else:
            for sheet, df in dfs.items():
                safe_sheet = str(sheet)[:31] if sheet else "Sheet"
                df.to_excel(writer, index=False, sheet_name=safe_sheet)
    return output.getvalue()


def render_export_buttons(df, base_filename, key_prefix, label_prefix="Baixar"):
    """Render side-by-side CSV and Excel download buttons for a DataFrame."""
    if df is None or df.empty:
        return
    col_csv, col_xlsx = st.columns(2)
    with col_csv:
        st.download_button(
            label=f"⬇️ {label_prefix} CSV",
            data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"{base_filename}.csv",
            mime="text/csv",
            key=f"{key_prefix}_csv",
            use_container_width=True,
        )
    with col_xlsx:
        st.download_button(
            label=f"⬇️ {label_prefix} Excel",
            data=to_excel_bytes(df),
            file_name=f"{base_filename}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}_xlsx",
            use_container_width=True,
        )


def get_img_as_base64(path):
    with open(path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()


def normalize_name(name):
    if pd.isna(name):
        return ""
    name = (
        unicodedata.normalize("NFKD", str(name))
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
        .strip()
    )
    return name


image_path = "imagem_fundo.png"
if os.path.exists(image_path):
    b64 = get_img_as_base64(image_path)
    st.markdown(
        f"""
        <style>
            .bg-top {{
                height: 40vh;  # ← MUDANÇA 1: Agora 33vh (terço superior)
                width: 100%;
                background-image: url(data:image/png;base64,{b64});
                background-size: cover;
                background-position: center center;
                background-repeat: no-repeat;
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown('<div class="bg-top"></div>', unsafe_allow_html=True)
st.markdown('<div style="margin-top: -600px;">', unsafe_allow_html=True)
st.markdown(
    """
<style>
</style>
""",
    unsafe_allow_html=True,
)


#st.markdown("""
#    <head>
#    <meta name="theme-color" content="#ffffff">
#    <link rel="manifest" href="manifest.json">
#    <link rel="icon" type="image/png" sizes="32x32" href="favicon-32x32.png">
#    <link rel="icon" type="image/png" sizes="16x16" href="favicon-16x16.png">
#    <link rel="apple-touch-icon" href="apple-touch-icon.png">
#</head>

#<style>
#  .titulo-principal {
#    font-family: 'Arial Black', Arial, sans-serif;
#    font-size: 3rem;
#    font-weight: bold;
#    color: #333333;
#    text-align: center;
#    margin-bottom: 1rem;
#    text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
#  }

#  .titulo-azul {
#    font-family: Arial, sans-serif;
#    font-size: 2.5rem;
#    color: #0066CC;
#    text-align: center;
#    margin-bottom: 0.8rem;
#    font-weight: 600;
#  }

#  .texto-detalhe {
#    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
#    font-size: 1.1rem;
#    color: #666666;
#    line-height: 1.7;
#    text-align: justify;
#    margin-bottom: 1rem;
#  }

#  .stExpander {
#    background-color: #f8f9fa;
#    border: 1px solid #dee2e6;
#    border-radius: 0.375rem;
#    padding: 1rem;
#    margin: 1rem 0;
#  }

#  .stExpander > div > label {
#    font-weight: 500;
#    color: #495057;
#  }
#</style>
#""", unsafe_allow_html=True)
def verificar_senha():
    if "senha_correta" not in st.session_state:
        st.session_state.senha_correta = False
    if not st.session_state.senha_correta:
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            caminho_logo = "logo_pbi.jpg"
            if os.path.exists(caminho_logo):
                st.image(caminho_logo, width=200)
        with st.form("login_form"):
            senha_digitada = st.text_input("Digite a senha para acessar:", type="password")
            entrar = st.form_submit_button("Acessar")
        if entrar:
            senha_correta = "Hapvida+Odonto"
            if senha_digitada == senha_correta:
                st.session_state.senha_correta = True
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")
        st.stop()
verificar_senha()

if "secao_ativa" not in st.session_state:
    st.session_state.secao_ativa = None
@st.cache_data
def carregar_dados():
    arquivo = "AUDITORIA_ODONTO_2026.xlsx"
    if not os.path.exists(arquivo):
        st.error(f"Arquivo '{arquivo}' não encontrado!")
        st.stop()
    def limpar_colunas(df):
        df.columns = df.columns.astype(str).str.strip()
        df.columns = df.columns.map(lambda x: unicodedata.normalize('NFKD', x).encode('ascii', 'ignore').decode('ascii'))
        df.columns = df.columns.str.lower().str.replace(r'[/\-\s]+', '_', regex=True)
        return df.fillna("-")
    glosas = pd.read_excel(arquivo, sheet_name="Glosas")
    procedimentos = pd.read_excel(arquivo, sheet_name="Procedimentos")
    regras_gerais = pd.read_excel(arquivo, sheet_name="Regras_Gerais")
    regras_espec = pd.read_excel(arquivo, sheet_name="Regras_Especialidade")
    produtos = pd.read_excel(arquivo, sheet_name="Produtos")
    glosas_limpo = limpar_colunas(glosas)
    proc_limpo = limpar_colunas(procedimentos)
    regras_gerais_limpo = limpar_colunas(regras_gerais)
    regras_espec_limpo = limpar_colunas(regras_espec)
    produtos_limpo = limpar_colunas(produtos)
    produtos_limpo = produtos_limpo.rename(columns={
    "produto": "codigo_do_produto",
    "descricao_completa": "nome_do_produto",
    "status": "status_do_produto",
    "procedimentos": "codigo_do_procedimento",
    "descricao_procedimento": "nome_do_procedimento",
    "grupo": "especialidade"
})
    def limpar_decimal(x):
        if pd.isna(x) or x == "-":
            return x
        try:
            return str(int(float(x)))
        except:
            return str(x)
    proc_limpo['codigo_interno'] = proc_limpo['codigo_interno'].apply(limpar_decimal)
    return glosas_limpo, proc_limpo, regras_gerais_limpo, regras_espec_limpo, produtos_limpo
    
def find_column(df, keywords):
    cols_lower = [col.lower() for col in df.columns]
    for kw in keywords:
        kw_lower = kw.lower()
        matches = get_close_matches(kw_lower, cols_lower, n=1, cutoff=0.6)
        if matches:
            idx = cols_lower.index(matches[0])
            return df.columns[idx]
    return None
 
# Carrega dados
glosas_df, proc_df, regras_gerais_df, regras_espec_df, produtos_df  = carregar_dados()
# Validações
colunas_glosas = ["n_da_glosa", "ativa", "descricao_interna", "tipo_de_glosa", "especialidade", "utilizacao", "subglosa", "como_evitar_a_glosa", "cabe_recurso", "como_recorrer", "justificativa", "origem_da_glosa"]
colunas_glosas_faltando = [c for c in colunas_glosas if c not in glosas_df.columns]
if colunas_glosas_faltando:
    st.error(f"Colunas faltando em Glosas: {colunas_glosas_faltando}")
    st.stop()
colunas_proc = ["codigo_interno", "tuss", "procedimento", "especialidade", "local_regiao", "procedimentos_pre_aprovados", "pre_requisito", "longevidade", "normas_tecnicas_e_observacoes"]
colunas_proc_faltando = [c for c in colunas_proc if c not in proc_df.columns]
if colunas_proc_faltando:
    st.error(f"Colunas faltando em Procedimentos: {colunas_proc_faltando}")
    st.stop()
# Logo e Título principal
col_logo_1, col_logo_2, col_logo_3 = st.columns([1, 1, 1])
with col_logo_2:
    if os.path.exists("logo_pbi.jpg"):
        st.image("logo_pbi.jpg", width=200)
#st.markdown('<div class="titulo-principal">AUDITORIA ODONTO - 2026</div>', unsafe_allow_html=True)
# TELA INICIAL: Apenas 2 botões (em colunas)
left_col, right_col = st.columns([1, 3])
with left_col:
    if st.button("🔍 TABELA DE PROCEDIMENTOS", key="btn_procedimentos", help="Consultar procedimentos e regras"):
        st.session_state.secao_ativa = "procedimentos"
        st.rerun()
    if st.button("📋 MANUAL DE GLOSAS", key="btn_glosas", help="Consultar glosas e detalhes"):
        st.session_state.secao_ativa = "glosas"
        st.rerun()
# Conteúdo condicional baseado na seção ativa
if st.session_state.secao_ativa == "glosas":
    st.markdown("### Manual de Glosas")
    glosas_df["label_busca"] = glosas_df["n_da_glosa"].astype(str) + " - " + glosas_df["descricao_interna"].astype(str)
    opcoes_glosa = sorted(glosas_df["label_busca"].unique())
    selecao_glosa = st.selectbox("Digite o número da glosa ou descrição:", [""] + opcoes_glosa, 
                                 format_func=lambda x: "Selecione..." if x == "" else x)


    if selecao_glosa:
        glosa_id = selecao_glosa.split(" - ")[0]
        dados_glosa = glosas_df[glosas_df["n_da_glosa"].astype(str) == glosa_id].copy()
        primeira = dados_glosa.iloc[0]
        st.markdown('<div class="titulo-azul">Detalhes da Glosa</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="texto-detalhe"><b>N DA GLOSA:</b> {primeira["n_da_glosa"]}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="texto-detalhe"><b>DESCRIÇÃO INTERNA:</b> {primeira["descricao_interna"]}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="texto-detalhe"><b>ORIGEM:</b> {primeira["origem_da_glosa"]}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="texto-detalhe"><b>ATIVA:</b> {primeira["ativa"]}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="texto-detalhe"><b>TIPO:</b> {primeira["tipo_de_glosa"]}</div>', unsafe_allow_html=True)
        st.markdown('<div class="titulo-azul">Especialidade / Subglosa</div>', unsafe_allow_html=True)
        for _, linha in dados_glosa.iterrows():
            with st.expander(f"{linha['especialidade']} - {linha['subglosa']}"):
                st.markdown(f"**ESPECIALIDADE:** {linha['especialidade']}")
                st.markdown(f"**UTILIZAÇÃO:** {linha['utilizacao']}")
                st.markdown(f"**SUBGLOSA:** {linha['subglosa']}")
                st.markdown(f"**COMO EVITAR:** {linha['como_evitar_a_glosa']}")
                st.markdown(f"**CABE RECURSO:** {linha['cabe_recurso']}")
                st.markdown(f"**COMO RECORRER:** {linha['como_recorrer']}")
                st.markdown(f"**JUSTIFICATIVA:** {linha['justificativa']}")

elif st.session_state.secao_ativa == "procedimentos":
    if "proc_select" not in st.session_state:
        st.session_state.proc_select = "Selecione um procedimento:"
    if "mostrar_cobertura" not in st.session_state:
        st.session_state.mostrar_cobertura = False
    if "detalhes_expandido" not in st.session_state:
        st.session_state.detalhes_expandido = False
    if "regras_gerais_expandido" not in st.session_state:
        st.session_state.regras_gerais_expandido = False
    if "regras_espec_expandido" not in st.session_state:
        st.session_state.regras_espec_expandido = False
    proc_df.columns = proc_df.columns.astype(str).str.strip()
    regras_gerais_df.columns = regras_gerais_df.columns.astype(str).str.strip()
    regras_espec_df.columns = regras_espec_df.columns.astype(str).str.strip()
    produtos_df.columns = produtos_df.columns.astype(str).str.strip()
    proc_df["label_busca"] = (
        proc_df["tuss"].astype(str) + " - " +
        proc_df["codigo_interno"].astype(str) + " - " +
        proc_df["procedimento"].astype(str)
    )
    opcoes_proc = [""] + sorted(proc_df["label_busca"].unique())
    col_busca, col_limpar = st.columns([4, 1])
    with col_busca:
        st.selectbox(
            "Selecione um procedimento:",
            opcoes_proc,
            key="proc_select"
        )
    row = None
    codigo_interno = None
    proc_row = None
    if st.session_state.proc_select!="Selecione um procedimento:":
        proc_row = proc_df[proc_df["label_busca"] == st.session_state.proc_select]
        if not proc_row.empty:
            row = proc_row.iloc[0]
            codigo_interno = str(row["codigo_interno"]).strip()
    if row is not None:
        with st.expander("Cobertura"):
            codigo_interno = str(row["codigo_interno"]).strip()
            produtos_df["codigo_do_procedimento"] = produtos_df["codigo_do_procedimento"].astype(str).str.strip()
            produtos_df["cobertura"] = produtos_df["cobertura"].astype(str).str.strip().str.lower()
            produtos_cobertos = produtos_df[
                (produtos_df["codigo_do_procedimento"] == codigo_interno) &
                (produtos_df["cobertura"] == "sim")
            ]
            if not produtos_cobertos.empty:
                st.write("Produtos cobertos:")
                for _, produto in produtos_cobertos.iterrows():
                    st.write(f"- {produto['nome_do_produto']} (Código: {produto['codigo_do_produto']})")
            else:
                st.write("Nenhum produto coberto encontrado para este procedimento.")
        with st.expander("Detalhes do Procedimento"):
            detalhes_df = proc_row.copy()
            nomes_personalizados = {
                "codigo_interno": "CÓDIGO INTERNO",
                "tuss": "TUSS",
                "procedimento": "PROCEDIMENTO",
                "especialidade": "ESPECIALIDADE",
                "local_regiao": "LOCAL / REGIÃO",
                "procedimentos_pre_aprovados": "PROCEDIMENTOS PRÉ-APROVADOS",
                "pre_requisito": "PRÉ-REQUISITOS",
                "longevidade": "LONGEVIDADE",
                "normas_tecnicas_e_observacoes": "NORMAS TÉCNICAS E OBSERVAÇÕES"
            }
            colunas_detalhes = [col for col in detalhes_df.columns if col not in ["label_busca"]]
            if not detalhes_df.empty:
                for col in colunas_detalhes:
                    valor = detalhes_df.iloc[0][col]
                    nome_exibicao = nomes_personalizados.get(col, col.upper())
                    st.write(f"**{nome_exibicao}**: {valor}")
            else:
                st.write("Nenhum detalhe encontrado para este procedimento.")
        with st.expander("Regras por Especialidade"):
            if st.session_state.proc_select:
                if "especialidade" in proc_df.columns and "especialidade" in regras_espec_df.columns and "regras_da_especialidade" in regras_espec_df.columns:
                    row = proc_row.iloc[0]
                    especialidade_proc = str(row["especialidade"]).strip().lower().replace(".0", "")
                    especialidades_regras = (
                        regras_espec_df["especialidade"]
                        .astype(str)
                        .str.strip()
                        .str.lower()
                        .str.replace(".0", "", regex=False)
                    )
                    match_regras = regras_espec_df[especialidades_regras == especialidade_proc]
                    if not match_regras.empty:
                        regras_texto = (
                            match_regras["regras_da_especialidade"]
                            .dropna()
                            .astype(str)
                            .str.strip()
                        )
                        regras_texto = regras_texto[regras_texto != ""]
                        if not regras_texto.empty:
                            for regra in regras_texto:
                                st.write(regra)
                        else:
                            st.write("Nenhuma regra específica encontrada para esta especialidade.")
                    else:
                        st.write("Nenhuma regra específica encontrada para esta especialidade.")
                else:
                    st.write(f"Colunas encontradas em Procedimentos: {list(proc_df.columns)}")
                    st.write(f"Colunas encontradas em Regras_Especialidade: {list(regras_espec_df.columns)}")
            else:
                st.info("Selecione um procedimento para visualizar as regras por especialidade.")
        with st.expander("Regras Gerais"):
            if st.session_state.proc_select:
                        regras_validas = regras_gerais_df["regras_gerais"].dropna().astype(str).str.strip()
                        regras_validas = regras_validas[regras_validas != ""]
                        if not regras_validas.empty:
                            for regra in regras_validas:
                                st.write(regra)
                        else:
                            st.write("Nenhuma regra geral encontrada na aba Regras_Gerais.")
            else:
                st.info("Selecione um procedimento para visualizar as regras gerais.")


st.set_page_config(
    page_title="Chat Auditoria Odontológica",
    page_icon="🪇",
    layout="wide"
)

#st.title("Faça uma pergunta ao Bob")

# Sidebar
#st.sidebar.title("Configurações")
#api_key = st.sidebar.text_input("Chave da API OpenAI", type="password")
#if st.sidebar.button("Limpar Chat"):
#    st.session_state.messages = []

# Carregar contexto do arquivo
#@st.cache_data
#def load_context():
#    filename = "Chat Auditoria Odontológica.txt"
#    encodings = ['utf-8', 'latin-1', 'cp-1252']
#    try:
#        with open(filename, 'rb') as f:
#            content_bytes = f.read()
#    except FileNotFoundError:
#        return "Arquivo não encontrado: Chat Auditoria Odontológica.txt"
#    except Exception as e:
#        return f"Erro ao abrir o arquivo: {str(e)}"
    
#    for encoding in encodings:
#        try:
#            content = content_bytes.decode(encoding)
#            return content
#        except UnicodeDecodeError:
#            continue
    
#    return "Erro ao carregar o arquivo: não foi possível decodificar com utf-8, latin-1 ou cp-1252."   

#context = load_context()
#if context.startswith("Erro") or context.startswith("Arquivo"):
#    st.error(f"❌ {context}")
#    st.stop()


# Histórico do chat
#if "messages" not in st.session_state:
#    st.session_state.messages = []

# Exibir mensagens do chat
#for message in st.session_state.messages:
#    with st.chat_message(message["role"]):
#        st.markdown(message["content"])

# Input do chat
#if prompt := st.chat_input("Digite sua pergunta sobre auditoria odontológica..."):
#    if not api_key:
#        st.error("Por favor, insira a chave da API OpenAI na sidebar.")
#        st.stop()
#
    # Adicionar mensagem do usuário
#    st.session_state.messages.append({"role": "user", "content": prompt})
#    with st.chat_message("user"):
#        st.markdown(prompt)

    # Resposta do assistente
#    with st.chat_message("assistant"):
#        message_placeholder = st.empty()
#        full_response = ""
#        try:
#            client = OpenAI(api_key=api_key)
#            response = client.chat.completions.create(
#                model="gpt-4o-mini",
#                messages=[
#                    {"role": "system", "content": f"Você é um especialista em auditoria odontológica. Use este contexto para responder às perguntas do usuário de forma precisa e útil:\n\n{context}"}
#                ] + st.session_state.messages,
#                temperature=0.7,
#            )
#            full_response = response.choices[0].message.content
#        except Exception as e:
#            full_response = f"Erro ao gerar resposta: {str(e)}"
#        message_placeholder.markdown(full_response)
#        st.session_state.messages.append({"role": "assistant", "content": full_response})    