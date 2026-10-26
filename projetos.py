import streamlit as st
import pandas as pd
from datetime import datetime, date
import os

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

EXCEL_FILE = "Projetos_sanear.xlsx"
LOGO_FILE = "docx_extracted/word/media/image1.jpeg"  # Caminho da imagem/logo do SANEAR

# Função para carregar e tratar os dados da planilha
def carregar_dados():
    try:
        df = pd.read_excel(EXCEL_FILE)
        df['INICIO'] = pd.to_datetime(df['INICIO']).dt.date
        df['PRAZO FINAL'] = pd.to_datetime(df['PRAZO FINAL']).dt.date
        df['DATA CONCLUSAO'] = pd.to_datetime(df['DATA CONCLUSAO']).dt.date
        return df
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo {EXCEL_FILE}: {e}")
        return pd.DataFrame(columns=["PROJETO", "Material/Serviço", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO"])

# Função para salvar alterações na planilha
def salvar_dados(df):
    try:
        df.to_excel(EXCEL_FILE, index=False)
        st.success("💧 Dados salvos com sucesso!")
    except Exception as e:
        st.error(f"Erro ao salvar alterações no arquivo: {e}")

# Função para calcular status e dias de atraso
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

# Estilo visual de cores do SANEAR para as tabelas
def estilar_status(val):
    if val == 'Em Atraso':
        return 'background-color: #ffebee; color: #c62828; font-weight: bold;'
    elif val == 'Concluído':
        return 'background-color: #e8f5e9; color: #2e7d32; font-weight: bold;'
    elif val == 'Em Andamento':
        return 'background-color: #e3f2fd; color: #1565c0; font-weight: bold;'
    return ''

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
    # Exibir Logotipo na Barra Lateral se existir
    if os.path.exists(LOGO_FILE):
        st.sidebar.image(LOGO_FILE, use_container_width=True)
    else:
        st.sidebar.markdown("## 💧 **SANEAR**")

    st.sidebar.title("🌊 Navegação")
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
        
        # Indicadores Chave
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📁 Total de Projetos", len(df_calculado))
        col2.metric("⏳ Em Andamento", len(df_calculado[df_calculado['STATUS'] == "Em Andamento"]))
        col3.metric("✅ Concluídos", len(df_calculado[df_calculado['STATUS'] == "Concluído"]))
        col4.metric("🚨 Em Atraso", len(df_calculado[df_calculado['STATUS'] == "Em Atraso"]))

        st.markdown("---")

        # Seção de Alerta de Projetos em Atraso
        df_atrasados = df_calculado[df_calculado['STATUS'] == "Em Atraso"].sort_values(by="DIAS DE ATRASO", ascending=False)
        st.subheader("🚨 Projetos Com Prazos Vencidos")
        
        if not df_atrasados.empty:
            df_atrasados_fmt = formatar_datas_para_exibicao(df_atrasados)
            st.dataframe(
                df_atrasados_fmt[["PROJETO", "RESPONSAVEL", "PRAZO FINAL", "DIAS DE ATRASO", "STATUS"]].style.map(estilar_status, subset=['STATUS']),
                use_container_width=True
            )
        else:
            st.success("🌿 Excelente! Nenhum projeto em atraso no momento.")

        st.markdown("---")

        # Tabela Geral
        st.subheader("📋 Acompanhamento de Prazos e Conclusão")
        df_geral_fmt = formatar_datas_para_exibicao(df_calculado)
        cols_ordem = ["PROJETO", "Material/Serviço", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO", "STATUS", "DIAS DE ATRASO"]
        
        st.dataframe(
            df_geral_fmt[cols_ordem].style.map(estilar_status, subset=['STATUS']),
            use_container_width=True
        )

    # 2. MÓDULO DE RELATÓRIOS
    elif opcao == "📋 Relatórios Ambientais":
        st.title("📋 Relatório Geral de Projetos e Serviços")

        st.subheader("🔍 Filtros de Consulta")
        col_f1, col_f2 = st.columns(2)

        responsaveis = ["Todos"] + list(df_calculado['RESPONSAVEL'].dropna().unique())
        resp_sel = col_f1.selectbox("👤 Responsável:", responsaveis)

        status_opcoes = ["Todos", "Em Atraso", "Em Andamento", "Concluído"]
        status_sel = col_f2.selectbox("📌 Status do Projeto:", status_opcoes)

        # Filtros
        df_filtrado = df_calculado.copy()
        if resp_sel != "Todos":
            df_filtrado = df_filtrado[df_filtrado['RESPONSAVEL'] == resp_sel]
        if status_sel != "Todos":
            df_filtrado = df_filtrado[df_filtrado['STATUS'] == status_sel]

        st.markdown("---")
        st.write(f"**Registros Localizados:** {len(df_filtrado)}")
        
        df_filtrado_fmt = formatar_datas_para_exibicao(df_filtrado)
        st.dataframe(
            df_filtrado_fmt.style.map(estilar_status, subset=['STATUS']),
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

    # 3. INSERÇÃO E EDIÇÃO DE DADOS
    elif opcao == "🛠️ Inserção e Gestão":
        st.title("🛠️ Cadastrar e Atualizar Projetos")

        aba_novo, aba_editar = st.tabs(["➕ Novo Projeto / Serviço", "✏️ Atualizar Status / Conclusão"])

        # Cadastro de novos projetos
        with aba_novo:
            st.subheader("Cadastrar Nova Demanda")
            with st.form("form_novo_projeto"):
                novo_nome = st.text_input("🏷️ Nome do Projeto / Obra")
                novo_tipo = st.selectbox("📦 Material / Serviço", ["Material", "Serviço", "Outro"])
                novo_resp = st.text_input("👤 Responsável Técnico")
                novo_inicio = st.date_input("📅 Data de Início", value=date.today(), format="DD/MM/YYYY")
                novo_prazo = st.date_input("⏳ Prazo Final Previsto", value=date.today(), format="DD/MM/YYYY")

                btn_cadastrar = st.form_submit_button("💾 Salvar Projeto")

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
                        salvar_dados(df_atualizado)
                        st.rerun()
                    else:
                        st.warning("⚠️ Preencha os campos obrigatórios (Nome do Projeto e Responsável).")

        # Atualização de Projetos
        with aba_editar:
            st.subheader("Registrar Conclusão de Projeto")
            if not df_bruto.empty:
                projeto_sel = st.selectbox("Selecione o Projeto:", df_bruto['PROJETO'].unique())
                idx = df_bruto[df_bruto['PROJETO'] == projeto_sel].index[0]

                data_atual_conc = df_bruto.loc[idx, 'DATA CONCLUSAO']
                val_inicial_conc = data_atual_conc if pd.notna(data_atual_conc) else date.today()

                nova_conclusao = st.date_input("✅ Data de Conclusão Efetiva:", value=val_inicial_conc, format="DD/MM/YYYY")

                if st.button("💾 Registrar Conclusão"):
                    df_bruto.loc[idx, 'DATA CONCLUSAO'] = nova_conclusao
                    cols_salvar = ["PROJETO", "Material/Serviço", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO"]
                    salvar_dados(df_bruto[cols_salvar])
                    st.rerun()
