import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
from streamlit_gsheets import GSheetsConnection

# Configuração da página
st.set_page_config(
    page_title="SANEAR - Gestão de Projetos",
    page_icon="💧",
    layout="wide"
)

# Estilização CSS customizada com a paleta oficial SANEAR (Azul e Verde)
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
    </style>
""", unsafe_allow_html=True)

LOGO_FILE = "docx_extracted/word/media/image1.jpeg"  # Caminho da imagem/logo do SANEAR

# --- CONEXÃO COM GOOGLE SHEETS ---
@st.cache_resource(ttl=60)
def obter_conexao_gsheets():
    return st.connection("gsheets", type=GSheetsConnection)

# Função para carregar os dados das abas
def carregar_dados_pac():
    conn = obter_conexao_gsheets()
    try:
        # Lê a aba NovoPac da planilha PROJETOS_SANEAR
        df = conn.read(worksheet="NovoPac", ttl="0s")
        df = df.dropna(how="all")  # Remove linhas totalmente vazias
        
        # Conversão de datas
        colunas_data = ['INICIO', 'PRAZO FINAL', 'DATA CONCLUSAO']
        for col in colunas_data:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
        return df
    except Exception as e:
        st.error(f"Erro ao carregar a aba 'NovoPac' no Google Sheets: {e}")
        return pd.DataFrame(columns=["PROJETO", "Material/Serviço", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO"])

def carregar_servidores():
    conn = obter_conexao_gsheets()
    try:
        # Lê a aba Servidor para obter a lista de responsáveis
        df_serv = conn.read(worksheet="Servidor", ttl="60s")
        df_serv = df_serv.dropna(how="all")
        if "NOME" in df_serv.columns:
            return df_serv["NOME"].dropna().unique().tolist()
        elif "RESPONSAVEL" in df_serv.columns:
            return df_serv["RESPONSAVEL"].dropna().unique().tolist()
        else:
            return df_serv.iloc[:, 0].dropna().unique().tolist()
    except Exception as e:
        st.warning(f"Não foi possível carregar a lista de servidores da aba 'Servidor': {e}")
        return []

# Função para atualizar a aba NovoPac
def salvar_dados_pac(df):
    conn = obter_conexao_gsheets()
    try:
        # Garante a formatação de data para texto simples antes de enviar
        df_salvar = df.copy()
        for col in ['INICIO', 'PRAZO FINAL', 'DATA CONCLUSAO']:
            if col in df_salvar.columns:
                df_salvar[col] = df_salvar[col].apply(
                    lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) and isinstance(x, (date, datetime)) else ""
                )
        conn.update(worksheet="NovoPac", data=df_salvar)
        st.success("💧 Dados atualizados no Google Sheets com sucesso!")
    except Exception as e:
        st.error(f"Erro ao salvar alterações no Google Sheets: {e}")

# Função para calcular status e dias de atraso
def calcular_status(df):
    hoje = date.today()
    status_list = []
    dias_atraso_list = []

    for index, row in df.iterrows():
        prazo = row.get('PRAZO FINAL', pd.NaT)
        conclusao = row.get('DATA CONCLUSAO', pd.NaT)

        if pd.notna(conclusao):
            status_list.append("Concluído")
            dias_atraso_list.append(0)
        elif pd.notna(prazo) and prazo < hoje:
            atraso = (hoje - prazo).days
            status_list.append("Em Atraso")
            dias_atraso_list.append(atraso)
        else:
            status_list.append("Em Andamento")
            dias_atraso_list.append(0)

    df['STATUS'] = status_list
    df['DIAS DE ATRASO'] = dias_atraso_list
    return df

# Formatação de data DD/MM/AAAA para exibição
def formatar_datas_para_exibicao(df):
    df_exibicao = df.copy()
    colunas_data = ['INICIO', 'PRAZO FINAL', 'DATA CONCLUSAO']
    for col in colunas_data:
        if col in df_exibicao.columns:
            df_exibicao[col] = df_exibicao[col].apply(
                lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) and isinstance(x, (date, datetime)) else (x if pd.notna(x) else "-")
            )
    return df_exibicao

# Estilo visual com cores de fundo para a tabela inteira (Atrasado vs Em Dia)
def estilar_linha_status(row):
    status = row.get('STATUS', '')
    if status == 'Em Atraso':
        # Cor de fundo vermelho suave para projetos em atraso
        return ['background-color: #ffcdd2; color: #b71c1c; font-weight: bold;'] * len(row)
    elif status in ['Em Andamento', 'Concluído']:
        # Cor de fundo verde suave para projetos em dia / concluídos
        return ['background-color: #c8e6c9; color: #1b5e20; font-weight: bold;'] * len(row)
    return [''] * len(row)

# --- MÓDULO DE AUTENTICAÇÃO ---
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

    st.sidebar.title("🌊 Navegação")
    opcao = st.sidebar.radio(
        "Selecione um módulo:",
        ["🏗️ Módulo Novo PAC", "📋 Relatórios Ambientais", "🛠️ Inserção e Gestão"]
    )
    
    if st.sidebar.button("🔒 Sair / Logout"):
        st.session_state["autenticado"] = False
        st.rerun()

    df_bruto = carregar_dados_pac()
    df_calculado = calcular_status(df_bruto)
    lista_servidores = carregar_servidores()

    # 1. DASHBOARD MÓDULO NOVO PAC
    if opcao == "🏗️ Módulo Novo PAC":
        st.title("🏗️ Módulo Novo PAC - Dashboard de Projetos")
        
        # Indicadores Chave
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📁 Total de Projetos", len(df_calculado))
        col2.metric("⏳ Em Andamento", len(df_calculado[df_calculado['STATUS'] == "Em Andamento"]))
        col3.metric("✅ Concluídos", len(df_calculado[df_calculado['STATUS'] == "Concluído"]))
        col4.metric("🚨 Em Atraso", len(df_calculado[df_calculado['STATUS'] == "Em Atraso"]))

        st.markdown("---")

        # Legenda das Cores
        st.markdown("""
        **Legenda de Destaque Visual:**
        * <span style="background-color: #ffcdd2; color: #b71c1c; padding: 4px 8px; border-radius: 4px; font-weight: bold;">Projetos em Atraso</span>
        * <span style="background-color: #c8e6c9; color: #1b5e20; padding: 4px 8px; border-radius: 4px; font-weight: bold;">Projetos Em Dia / Concluídos</span>
        """, unsafe_allow_html=True)
        st.write("")

        # Tabela Geral com Destacamento de Cores
        st.subheader("📋 Acompanhamento de Projetos Novo PAC")
        df_geral_fmt = formatar_datas_para_exibicao(df_calculado)
        cols_ordem = ["PROJETO", "Material/Serviço", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO", "STATUS", "DIAS DE ATRASO"]
        
        cols_existentes = [col for col in cols_ordem if col in df_geral_fmt.columns]
        
        st.dataframe(
            df_geral_fmt[cols_existentes].style.apply(estilar_linha_status, axis=1),
            use_container_width=True
        )

    # 2. MÓDULO DE RELATÓRIOS
    elif opcao == "📋 Relatórios Ambientais":
        st.title("📋 Relatório Geral de Projetos e Serviços (Novo PAC)")

        st.subheader("🔍 Filtros de Consulta")
        col_f1, col_f2 = st.columns(2)

        responsaveis = ["Todos"] + list(df_calculado['RESPONSAVEL'].dropna().unique()) if 'RESPONSAVEL' in df_calculado.columns else ["Todos"]
        resp_sel = col_f1.selectbox("👤 Responsável:", responsaveis)

        status_opcoes = ["Todos", "Em Atraso", "Em Andamento", "Concluído"]
        status_sel = col_f2.selectbox("📌 Status do Projeto:", status_opcoes)

        # Aplicação dos filtros
        df_filtrado = df_calculado.copy()
        if resp_sel != "Todos" and 'RESPONSAVEL' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['RESPONSAVEL'] == resp_sel]
        if status_sel != "Todos":
            df_filtrado = df_filtrado[df_filtrado['STATUS'] == status_sel]

        st.markdown("---")
        st.write(f"**Registros Localizados:** {len(df_filtrado)}")
        
        df_filtrado_fmt = formatar_datas_para_exibicao(df_filtrado)
        st.dataframe(
            df_filtrado_fmt.style.apply(estilar_linha_status, axis=1),
            use_container_width=True
        )

        # Download do relatório em CSV
        csv = df_filtrado_fmt.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Exportar Relatório em CSV",
            data=csv,
            file_name="relatorio_projetos_novo_pac.csv",
            mime="text/csv"
        )

    # 3. INSERÇÃO E EDIÇÃO DE DADOS
    elif opcao == "🛠️ Inserção e Gestão":
        st.title("🛠️ Cadastrar e Atualizar Projetos (Novo PAC)")

        aba_novo, aba_editar = st.tabs(["➕ Novo Projeto / Serviço", "✏️ Atualizar Status / Conclusão"])

        # Cadastro de novos projetos
        with aba_novo:
            st.subheader("Cadastrar Novo Projeto na Planilha NovoPac")
            with st.form("form_novo_projeto"):
                novo_nome = st.text_input("🏷️ Nome do Projeto / Obra")
                novo_tipo = st.selectbox("📦 Material / Serviço", ["Material", "Serviço", "Outro"])
                
                # Seleção do Responsável buscando da planilha Servidor
                if lista_servidores:
                    novo_resp = st.selectbox("👤 Responsável Técnico (Aba Servidor)", lista_servidores)
                else:
                    novo_resp = st.text_input("👤 Responsável Técnico")

                novo_inicio = st.date_input("📅 Data de Início", value=date.today(), format="DD/MM/YYYY")
                novo_prazo = st.date_input("⏳ Prazo Final Previsto", value=date.today(), format="DD/MM/YYYY")

                btn_cadastrar = st.form_submit_button("💾 Salvar Projeto no Google Sheets")

                if btn_cadastrar:
                    if novo_nome and novo_resp:
                        novo_registro = pd.DataFrame([{
                            "PROJETO": novo_nome,
                            "Material/Serviço": novo_tipo,
                            "RESPONSAVEL": novo_resp,
                            "INICIO": novo_inicio,
                            "PRAZO FINAL": novo_prazo,
                            "DATA CONCLUSAO": None
                        }])
                        cols_base = ["PROJETO", "Material/Serviço", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO"]
                        df_atualizado = pd.concat([df_bruto[cols_base], novo_registro], ignore_index=True)
                        salvar_dados_pac(df_atualizado)
                        st.rerun()
                    else:
                        st.warning("⚠️ Preencha os campos obrigatórios (Nome do Projeto e Responsável).")

        # Atualização de Projetos
        with aba_editar:
            st.subheader("Registrar Conclusão / Atualizar Data")
            if not df_bruto.empty and "PROJETO" in df_bruto.columns:
                projeto_sel = st.selectbox("Selecione o Projeto:", df_bruto['PROJETO'].unique())
                idx = df_bruto[df_bruto['PROJETO'] == projeto_sel].index[0]

                data_atual_conc = df_bruto.loc[idx, 'DATA CONCLUSAO']
                val_inicial_conc = data_atual_conc if pd.notna(data_atual_conc) and isinstance(data_atual_conc, date) else date.today()

                nova_conclusao = st.date_input("✅ Data de Conclusão Efetiva:", value=val_inicial_conc, format="DD/MM/YYYY")

                if st.button("💾 Salvar Conclusão no Google Sheets"):
                    df_bruto.loc[idx, 'DATA CONCLUSAO'] = nova_conclusao
                    cols_salvar = ["PROJETO", "Material/Serviço", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO"]
                    salvar_dados_pac(df_bruto[cols_salvar])
                    st.rerun()
