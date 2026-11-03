import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import os
import gspread

# Configuração da página
st.set_page_config(
    page_title="SANEAR - Gestão de Projetos",
    page_icon="💧",
    layout="wide"
)

# Estilização CSS customizada com a paleta oficial SANEAR e Cartões Personalizados
st.markdown("""
    <style>
    /* Cor dos botões principais */
    div.stButton > button:first-child {
        background-color: #0082cb;
        color: white;
        border-radius: 8px;
        border: none;
        font-weight: bold;
    }
    div.stButton > button:first-child:hover {
        background-color: #014f86;
        color: white;
    }
    /* Estilo dos títulos e destaques */
    .stApp h1 {
        color: #014f86;
    }
    .stApp h2, .stApp h3 {
        color: #2ba84a;
    }
    /* Borda da barra lateral */
    section[data-testid="stSidebar"] {
        background-color: #f4f9fc;
        border-right: 2px solid #e0f2fe;
    }
    
    /* Estilos dos Cartões de Métricas (Dashboard) */
    .metric-card {
        border-radius: 10px;
        padding: 15px;
        color: white;
        text-align: center;
        font-weight: bold;
        margin-bottom: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .metric-card h3 {
        color: white !important;
        margin: 0;
        font-size: 1rem;
    }
    .metric-card p {
        font-size: 1.8rem;
        margin: 5px 0 0 0;
    }
    .card-total { background-color: #0082cb; }
    .card-verde { background-color: #2e7d32; }
    .card-amarelo { background-color: #fbc02d; color: #333 !important; }
    .card-amarelo h3 { color: #333 !important; }
    .card-vermelho { background-color: #c62828; }
    </style>
""", unsafe_allow_html=True)

LOGO_FILE = "docx_extracted/word/media/image1.jpeg"
NOME_PLANILHA_GOOGLE = "PROJETOS_SANEAR"

# Colunas padrão do sistema
COLUNAS_PADRAO = ["PROJETO", "TIPO", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO"]

# Função para conectar ao Google Sheets
@st.cache_resource(ttl=300)
def conectar_google_sheets():
    try:
        # Tenta carregar credenciais das Secrets do Streamlit ou arquivo local
        if "gcp_service_account" in st.secrets:
            gc = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        else:
            gc = gspread.service_account(filename="credentials.json")
        return gc
    except Exception as e:
        st.error(f"Erro ao conectar com o Google Sheets: {e}")
        return None

# Função para carregar e tratar os dados do Google Sheets
def carregar_dados():
    gc = conectar_google_sheets()
    if not gc:
        return pd.DataFrame(columns=COLUNAS_PADRAO)
    
    try:
        sh = gc.open(NOME_PLANILHA_GOOGLE)
        
        # 1. Carrega a aba NovoPac
        wks_pac = sh.worksheet("NovoPac")
        dados_pac = wks_pac.get_all_records()
        df_pac = pd.DataFrame(dados_pac)
        df_pac.columns = df_pac.columns.str.strip().str.upper()

        # 2. Carrega a aba Servidor para obter responsáveis
        try:
            wks_servidor = sh.worksheet("Servidor")
            dados_servidor = wks_servidor.get_all_records()
            df_servidor = pd.DataFrame(dados_servidor)
            df_servidor.columns = df_servidor.columns.str.strip().str.upper()
        except Exception:
            df_servidor = pd.DataFrame()

        # Garante a existência das colunas
        for col in COLUNAS_PADRAO:
            if col not in df_pac.columns:
                df_pac[col] = None

        # Se houver relação de responsáveis na aba Servidor, mescla ou atualiza
        if not df_servidor.empty and "PROJETO" in df_servidor.columns and "RESPONSAVEL" in df_servidor.columns:
            df_pac = df_pac.drop(columns=["RESPONSAVEL"], errors='ignore')
            df_pac = df_pac.merge(df_servidor[["PROJETO", "RESPONSAVEL"]], on="PROJETO", how="left")

        # Tratamento de datas
        for col_data in ['INICIO', 'PRAZO FINAL', 'DATA CONCLUSAO']:
            if col_data in df_pac.columns:
                df_pac[col_data] = pd.to_datetime(df_pac[col_data], errors='coerce').dt.date

        return df_pac[COLUNAS_PADRAO]

    except Exception as e:
        st.error(f"Erro ao ler as abas 'NovoPac' / 'Servidor' no Google Sheets: {e}")
        return pd.DataFrame(columns=COLUNAS_PADRAO)

# Função para salvar novo projeto na aba NovoPac
def salvar_dados_google(novo_df_row):
    gc = conectar_google_sheets()
    if gc:
        try:
            sh = gc.open(NOME_PLANILHA_GOOGLE)
            wks = sh.worksheet("NovoPac")
            wks.append_row(novo_df_row.values.tolist()[0])
            st.success("💧 Dados salvos no Google Sheets com sucesso!")
            st.cache_resource.clear()
        except Exception as e:
            st.error(f"Erro ao salvar no Google Sheets: {e}")

# Função para calcular status, dias de atraso e status do dashboard
def calcular_status(df):
    hoje = date.today()
    status_list = []
    dias_atraso_list = []

    for index, row in df.iterrows():
        prazo = row['PRAZO FINAL']
        conclusao = row['DATA CONCLUSAO']

        if pd.notna(conclusao):
            status_list.append("Concluído")
            dias_atraso_list.append(0)
        elif pd.notna(prazo):
            if prazo < hoje:
                atraso = (hoje - prazo).days
                status_list.append("Em Atraso")
                dias_atraso_list.append(atraso)
            elif prazo <= hoje + timedelta(days=7):
                status_list.append("Próximo do Vencimento")
                dias_atraso_list.append(0)
            else:
                status_list.append("No Prazo")
                dias_atraso_list.append(0)
        else:
            status_list.append("No Prazo")
            dias_atraso_list.append(0)

    df['STATUS'] = status_list
    df['DIAS DE ATRASO'] = dias_atraso_list
    return df

# Formatação de data DD/MM/AAAA
def formatar_datas_para_exibicao(df):
    df_exibicao = df.copy()
    colunas_data = ['INICIO', 'PRAZO FINAL', 'DATA CONCLUSAO']
    for col in colunas_data:
        if col in df_exibicao.columns:
            df_exibicao[col] = df_exibicao[col].apply(
                lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) and isinstance(x, (date, datetime)) else (x if pd.notna(x) else "-")
            )
    return df_exibicao

# Função para estilar a LINHA INTEIRA na tabela
def estilar_linha_inteira(row):
    status = row['STATUS']
    if status == 'Em Atraso':
        # Linha inteira em vermelho suave
        return ['background-color: #ffcdd2; color: #b71c1c; font-weight: bold;'] * len(row)
    elif status in ['No Prazo', 'Concluído']:
        # Linha inteira em verde suave
        return ['background-color: #c8e6c9; color: #1b5e20; font-weight: bold;'] * len(row)
    elif status == 'Próximo do Vencimento':
        # Linha inteira em amarelo suave
        return ['background-color: #fff9c4; color: #f57f17; font-weight: bold;'] * len(row)
    return [''] * len(row)

# Autenticação do Usuário
def autenticacao():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if not st.session_state["autenticado"]:
        st.sidebar.title("🛡️ Acesso Restrito")
        senha_correta = "sanear123"
        senha_digitada = st.sidebar.text_input("🔑 Digite a senha de acesso:", type="password")

        if st.sidebar.button("🔓 Entrar no Sistema"):
            if senha_digitada == senha_correta:
                st.session_state["autenticado"] = True
                st.sidebar.success("Acesso liberado!")
                st.rerun()
            else:
                st.sidebar.error("Senha incorreta!")
        return False
    return True

# --- APLICAÇÃO PRINCIPAL ---
if autenticacao():
    if os.path.exists(LOGO_FILE):
        st.sidebar.image(LOGO_FILE, use_container_width=True)
    else:
        st.sidebar.markdown("## 💧 **SANEAR**")

    opcao = st.sidebar.radio(
        "Selecione um módulo:",
        ["📊 Dashboard & Prazos", "📋 Relatórios Ambientais", "🛠️ Inserção e Gestão"]
    )
    
    if st.sidebar.button("🔒 Sair / Logout"):
        st.session_state["autenticado"] = False
        st.rerun()

    df_bruto = carregar_dados()
    df_calculado = calcular_status(df_bruto)

    # 1. DASHBOARD DE APRESENTAÇÃO
    if opcao == "📊 Dashboard & Prazos":
        st.title("💧 Painel de Gestão de Projetos - SANEAR")
        
        # Contadores de Status
        total_proj = len(df_calculado)
        no_prazo = len(df_calculado[df_calculado['STATUS'].isin(["No Prazo", "Concluído"])])
        prox_venc = len(df_calculado[df_calculado['STATUS'] == "Próximo do Vencimento"])
        em_atraso = len(df_calculado[df_calculado['STATUS'] == "Em Atraso"])

        # Quadrados Coloridos do Dashboard
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f"""<div class="metric-card card-total"><h3>📁 Total de Projetos</h3><p>{total_proj}</p></div>""", unsafe_allow_html=True)
        with col2:
            st.markdown(f"""<div class="metric-card card-verde"><h3>✅ Dentro do Prazo</h3><p>{no_prazo}</p></div>""", unsafe_allow_html=True)
        with col3:
            st.markdown(f"""<div class="metric-card card-amarelo"><h3>⚠️ Próximo do Vencimento</h3><p>{prox_venc}</p></div>""", unsafe_allow_html=True)
        with col4:
            st.markdown(f"""<div class="metric-card card-vermelho"><h3>🚨 Em Atraso</h3><p>{em_atraso}</p></div>""", unsafe_allow_html=True)

        st.markdown("---")

        # Seção de Alerta de Projetos em Atraso
        df_atrasados = df_calculado[df_calculado['STATUS'] == "Em Atraso"].sort_values(by="DIAS DE ATRASO", ascending=False)
        st.subheader("🚨 Projetos Com Prazos Vencidos")
        
        if not df_atrasados.empty:
            df_atrasados_fmt = formatar_datas_para_exibicao(df_atrasados)
            st.dataframe(
                df_atrasados_fmt[["PROJETO", "TIPO", "RESPONSAVEL", "PRAZO FINAL", "DIAS DE ATRASO", "STATUS"]].style.apply(estilar_linha_inteira, axis=1),
                use_container_width=True
            )
        else:
            st.success("🌿 Excelente! Nenhum projeto em atraso no momento.")

        st.markdown("---")

        # Tabela Geral com Destaque na Linha Inteira
        st.subheader("📋 Acompanhamento Geral de Prazos")
        df_geral_fmt = formatar_datas_para_exibicao(df_calculado)
        cols_ordem = ["PROJETO", "TIPO", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO", "STATUS", "DIAS DE ATRASO"]
        
        st.dataframe(
            df_geral_fmt[cols_ordem].style.apply(estilar_linha_inteira, axis=1),
            use_container_width=True
        )

    # 2. MÓDULO DE RELATÓRIOS
    elif opcao == "📋 Relatórios Ambientais":
        st.title("📋 Relatório Geral de Projetos e Serviços")

        st.subheader("🔍 Filtros de Consulta")
        col_f1, col_f2 = st.columns(2)

        responsaveis = ["Todos"] + list(df_calculado['RESPONSAVEL'].dropna().unique())
        resp_sel = col_f1.selectbox("👤 Responsável:", responsaveis)

        status_opcoes = ["Todos", "Em Atraso", "Próximo do Vencimento", "No Prazo", "Concluído"]
        status_sel = col_f2.selectbox("📌 Status do Projeto:", status_opcoes)

        # Aplicação de Filtros
        df_filtrado = df_calculado.copy()
        if resp_sel != "Todos":
            df_filtrado = df_filtrado[df_filtrado['RESPONSAVEL'] == resp_sel]
        if status_sel != "Todos":
            df_filtrado = df_filtrado[df_filtrado['STATUS'] == status_sel]

        st.markdown("---")
        st.write(f"**Registros Localizados:** {len(df_filtrado)}")
        
        df_filtrado_fmt = formatar_datas_para_exibicao(df_filtrado)
        st.dataframe(
            df_filtrado_fmt.style.apply(estilar_linha_inteira, axis=1),
            use_container_width=True
        )

        # Download do relatório em CSV
        csv = df_filtrado_fmt.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Exportar Relatório em CSV",
            data=csv,
            file_name="relatorio_projetos_sanear.csv",
            mime="text/csv"
        )

    # 3. INSERÇÃO E GESTÃO
    elif opcao == "🛠️ Inserção e Gestão":
        st.title("🛠️ Cadastrar Demanda no Google Sheets")

        with st.form("form_novo_projeto"):
            novo_nome = st.text_input("🏷️ Nome do Projeto / Obra")
            novo_tipo = st.selectbox("📦 Tipo", ["Material", "Serviço", "Outro"])
            novo_resp = st.text_input("👤 Responsável Técnico")
            novo_inicio = st.date_input("📅 Data de Início", value=date.today(), format="DD/MM/YYYY")
            novo_prazo = st.date_input("⏳ Prazo Final Previsto", value=date.today(), format="DD/MM/YYYY")

            btn_cadastrar = st.form_submit_button("💾 Salvar na Planilha NovoPac")

            if btn_cadastrar:
                if novo_nome and novo_resp:
                    novo_registro = pd.DataFrame([{
                        "PROJETO": novo_nome,
                        "TIPO": novo_tipo,
                        "RESPONSAVEL": novo_resp,
                        "INICIO": novo_inicio.strftime('%Y-%m-%d'),
                        "PRAZO FINAL": novo_prazo.strftime('%Y-%m-%d'),
                        "DATA CONCLUSAO": ""
                    }])
                    salvar_dados_google(novo_registro)
                    st.rerun()
                else:
                    st.warning("⚠️ Preencha os campos obrigatórios (Nome do Projeto e Responsável).")
