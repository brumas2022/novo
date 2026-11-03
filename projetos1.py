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

# Estilização CSS customizada (Paleta SANEAR + Estilos dos Cards do Dashboard)
st.markdown("""
    <style>
    /* Botões principais */
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
    .stApp h1 { color: #014f86; }
    .stApp h2, .stApp h3 { color: #2ba84a; }
    
    /* Barra lateral */
    section[data-testid="stSidebar"] {
        background-color: #f4f9fc;
        border-right: 2px solid #e0f2fe;
    }

    /* Cards coloridos para o Dashboard */
    .card-dashboard {
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
        font-weight: bold;
        margin-bottom: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .card-total { background-color: #0082cb; }
    .card-verde { background-color: #2e7d32; }
    .card-amarelo { background-color: #fbc02d; color: #333333; }
    .card-vermelho { background-color: #c62828; }
    .card-num { font-size: 2.2rem; margin-top: 5px; }
    .card-title { font-size: 1.1rem; }
    </style>
""", unsafe_allow_html=True)

LOGO_FILE = "docx_extracted/word/media/image1.jpeg"

# Inicializa conexão GSheets
def obter_conexao():
    return st.connection("gsheets", type=GSheetsConnection)

# Função para conectar e carregar os dados do Google Sheets
def carregar_dados_gsheets():
    try:
        conn = obter_conexao()
        
        # Lê a aba NovoPac da planilha PROJETOS_SANEAR
        df_projetos = conn.read(worksheet="NovoPac", ttl="0")
        
        # Lê a aba Servidor da planilha PROJETOS_SANEAR
        df_servidor = conn.read(worksheet="Servidor", ttl="0")

        # Limpeza básica nos nomes de colunas
        df_projetos.columns = df_projetos.columns.astype(str).str.strip()
        df_servidor.columns = df_servidor.columns.astype(str).str.strip()

        # Garante que a coluna de responsável exista (cruza com aba Servidor se disponível)
        if "RESPONSAVEL" in df_servidor.columns and "RESPONSAVEL" not in df_projetos.columns:
            if "PROJETO" in df_servidor.columns and "PROJETO" in df_projetos.columns:
                df_projetos = df_projetos.merge(df_servidor[["PROJETO", "RESPONSAVEL"]], on="PROJETO", how="left")
            else:
                df_projetos["RESPONSAVEL"] = df_servidor["RESPONSAVEL"]
        elif "RESPONSAVEL" not in df_projetos.columns:
            df_projetos["RESPONSAVEL"] = "Não Definido"

        # Colunas obrigatórias
        cols_obrigatorias = ["PROJETO", "TIPO", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO"]
        for col in cols_obrigatorias:
            if col not in df_projetos.columns:
                df_projetos[col] = None

        # Tratamento de datas
        for col_data in ['INICIO', 'PRAZO FINAL', 'DATA CONCLUSAO']:
            df_projetos[col_data] = pd.to_datetime(df_projetos[col_data], errors='coerce').dt.date

        return df_projetos[cols_obrigatorias], df_servidor

    except Exception as e:
        st.error(f"⚠️ Erro ao conectar ao Google Sheets via secrets: {e}")
        cols = ["PROJETO", "TIPO", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO"]
        return pd.DataFrame(columns=cols), pd.DataFrame()

# Função para atualizar a aba NovoPac no Google Sheets
def salvar_dados_gsheets(df):
    try:
        conn = obter_conexao()
        df_salvar = df[["PROJETO", "TIPO", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO"]].copy()
        
        # Converte datas para string no formato AAAA-MM-DD para salvar corretamente no Sheets
        for col in ['INICIO', 'PRAZO FINAL', 'DATA CONCLUSAO']:
            df_salvar[col] = df_salvar[col].apply(lambda x: str(x) if pd.notna(x) and x != "" else "")
            
        conn.update(worksheet="NovoPac", data=df_salvar)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"❌ Erro ao salvar dados no Google Sheets: {e}")
        return False

# Função para calcular o status e regras de vencimento
def calcular_status(df):
    hoje = date.today()
    status_list = []
    dias_atraso_list = []
    cor_linha_list = []

    for index, row in df.iterrows():
        prazo = row['PRAZO FINAL']
        conclusao = row['DATA CONCLUSAO']

        if pd.notna(conclusao) and str(conclusao).strip() != "" and str(conclusao).strip() != "None":
            status_list.append("Concluído")
            dias_atraso_list.append(0)
            cor_linha_list.append("verde")
        elif pd.notna(prazo) and str(prazo).strip() != "" and str(prazo).strip() != "None":
            dias_para_vencer = (prazo - hoje).days
            if prazo < hoje:
                atraso = (hoje - prazo).days
                status_list.append("Em Atraso")
                dias_atraso_list.append(atraso)
                cor_linha_list.append("vermelho")
            elif 0 <= dias_para_vencer <= 7:
                status_list.append("Próximo do Vencimento")
                dias_atraso_list.append(0)
                cor_linha_list.append("amarelo")
            else:
                status_list.append("No Prazo")
                dias_atraso_list.append(0)
                cor_linha_list.append("verde")
        else:
            status_list.append("Sem Prazo")
            dias_atraso_list.append(0)
            cor_linha_list.append("verde")

    df['STATUS'] = status_list
    df['DIAS DE ATRASO'] = dias_atraso_list
    df['COR_LINHA'] = cor_linha_list
    return df

# Função para destacar a LINHA INTEIRA na tabela de acordo com o status
def estilar_linha_inteira(row):
    cor = row.get('COR_LINHA', '')
    if cor == 'vermelho':
        return ['background-color: #ffcdd2; color: #b71c1c; font-weight: bold;'] * len(row)
    elif cor == 'amarelo':
        return ['background-color: #fff9c4; color: #f57f17; font-weight: bold;'] * len(row)
    elif cor == 'verde':
        return ['background-color: #c8e6c9; color: #1b5e20; font-weight: bold;'] * len(row)
    return [''] * len(row)

# Formatação de exibição das datas (DD/MM/AAAA)
def formatar_datas_para_exibicao(df):
    df_exibicao = df.copy()
    colunas_data = ['INICIO', 'PRAZO FINAL', 'DATA CONCLUSAO']
    for col in colunas_data:
        if col in df_exibicao.columns:
            df_exibicao[col] = df_exibicao[col].apply(
                lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) and isinstance(x, (date, datetime)) else "-"
            )
    return df_exibicao

# Autenticação Simples
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
        [
            "📊 Dashboard & Prazos", 
            "🛠️ Gestão de Projetos (Incluir / Editar / Remover)",
            "📋 Relatórios Ambientais"
        ]
    )
    
    if st.sidebar.button("🔒 Sair / Logout"):
        st.session_state["autenticado"] = False
        st.rerun()

    # Carrega os dados diretamente do Google Sheets
    df_bruto, df_servidor = carregar_dados_gsheets()
    df_calculado = calcular_status(df_bruto.copy())

    # ---------------------------------------------------------
    # 1. DASHBOARD & PRAZOS
    # ---------------------------------------------------------
    if opcao == "📊 Dashboard & Prazos":
        st.title("💧 Painel de Gestão de Projetos - SANEAR")
        
        # Métricas calculadas
        total_projetos = len(df_calculado)
        no_prazo = len(df_calculado[df_calculado['COR_LINHA'] == "verde"])
        prox_vencimento = len(df_calculado[df_calculado['COR_LINHA'] == "amarelo"])
        em_atraso = len(df_calculado[df_calculado['COR_LINHA'] == "vermelho"])

        # Quadrados / Cards customizados
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""
                <div class="card-dashboard card-total">
                    <div class="card-title">📁 Total de Projetos</div>
                    <div class="card-num">{total_projetos}</div>
                </div>
            """, unsafe_allow_html=True)

        with c2:
            st.markdown(f"""
                <div class="card-dashboard card-verde">
                    <div class="card-title">✅ No Prazo / Concluídos</div>
                    <div class="card-num">{no_prazo}</div>
                </div>
            """, unsafe_allow_html=True)

        with c3:
            st.markdown(f"""
                <div class="card-dashboard card-amarelo">
                    <div class="card-title">⚠️ Próximo Vencimento (7d)</div>
                    <div class="card-num">{prox_vencimento}</div>
                </div>
            """, unsafe_allow_html=True)

        with c4:
            st.markdown(f"""
                <div class="card-dashboard card-vermelho">
                    <div class="card-title">🚨 Em Atraso</div>
                    <div class="card-num">{em_atraso}</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("---")

        # Tabela Geral com Destaque por Cores na Linha Inteira
        st.subheader("📋 Acompanhamento Geral de Todos os Projetos")
        cols_exibicao = ["PROJETO", "TIPO", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO", "STATUS", "DIAS DE ATRASO"]
        
        if not df_calculado.empty:
            df_geral_fmt = formatar_datas_para_exibicao(df_calculado)
            st.dataframe(
                df_geral_fmt[cols_exibicao].style.apply(estilar_linha_inteira, axis=1),
                use_container_width=True
            )
        else:
            st.info("Nenhum projeto cadastrado no momento.")

    # ---------------------------------------------------------
    # 2. MÓDULO DE INCLUSÃO, EDIÇÃO E REMOÇÃO (CRUD)
    # ---------------------------------------------------------
    elif opcao == "🛠️ Gestão de Projetos (Incluir / Editar / Remover)":
        st.title("🛠️ Gestão e Manipulação de Projetos")

        aba_novo, aba_editar, aba_remover = st.tabs([
            "➕ Incluir Novo Projeto", 
            "✏️ Editar / Concluir Projeto", 
            "🗑️ Remover Projeto"
        ])

        # Obter lista de responsáveis vindos da aba Servidor (se disponível)
        lista_servidores = []
        if not df_servidor.empty and "RESPONSAVEL" in df_servidor.columns:
            lista_servidores = list(df_servidor["RESPONSAVEL"].dropna().unique())

        # ABA 1: INCLUSÃO DE PROJETO
        with aba_novo:
            st.subheader("➕ Cadastrar Novo Projeto")
            with st.form("form_novo_projeto", clear_on_submit=True):
                novo_nome = st.text_input("🏷️ Nome do Projeto / Obra")
                novo_tipo = st.selectbox("📦 Tipo", ["Material", "Serviço", "Obra", "Outro"])
                
                if lista_servidores:
                    novo_resp = st.selectbox("👤 Responsável Técnico (Aba Servidor)", lista_servidores)
                else:
                    novo_resp = st.text_input("👤 Responsável Técnico")

                col_dt1, col_dt2 = st.columns(2)
                novo_inicio = col_dt1.date_input("📅 Data de Início", value=date.today(), format="DD/MM/YYYY")
                novo_prazo = col_dt2.date_input("⏳ Prazo Final Previsto", value=date.today(), format="DD/MM/YYYY")

                btn_cadastrar = st.form_submit_button("💾 Salvar Novo Projeto")

                if btn_cadastrar:
                    if novo_nome and novo_resp:
                        novo_registro = pd.DataFrame([{
                            "PROJETO": novo_nome,
                            "TIPO": novo_tipo,
                            "RESPONSAVEL": novo_resp,
                            "INICIO": novo_inicio,
                            "PRAZO FINAL": novo_prazo,
                            "DATA CONCLUSAO": None
                        }])
                        df_atualizado = pd.concat([df_bruto, novo_registro], ignore_index=True)
                        if salvar_dados_gsheets(df_atualizado):
                            st.success(f"✅ Projeto '{novo_nome}' inserido com sucesso!")
                            st.rerun()
                    else:
                        st.warning("⚠️ Preencha os campos obrigatórios (Nome do Projeto e Responsável).")

        # ABA 2: EDIÇÃO DE PROJETO
        with aba_editar:
            st.subheader("✏️ Editar ou Registrar Conclusão")
            if not df_bruto.empty:
                projetos_disponiveis = df_bruto["PROJETO"].dropna().unique()
                projeto_sel = st.selectbox("Selecione o Projeto para Editar:", projetos_disponiveis, key="select_editar")
                
                idx = df_bruto[df_bruto['PROJETO'] == projeto_sel].index[0]
                row_atual = df_bruto.loc[idx]

                with st.form("form_editar_projeto"):
                    edit_nome = st.text_input("🏷️ Nome do Projeto", value=str(row_atual["PROJETO"]))
                    edit_tipo = st.text_input("📦 Tipo", value=str(row_atual["TIPO"]) if pd.notna(row_atual["TIPO"]) else "Serviço")
                    
                    if lista_servidores:
                        val_resp_idx = lista_servidores.index(row_atual["RESPONSAVEL"]) if row_atual["RESPONSAVEL"] in lista_servidores else 0
                        edit_resp = st.selectbox("👤 Responsável", lista_servidores, index=val_resp_idx)
                    else:
                        edit_resp = st.text_input("👤 Responsável", value=str(row_atual["RESPONSAVEL"]))

                    col_ed1, col_ed2, col_ed3 = st.columns(3)
                    
                    dt_inc = row_atual["INICIO"] if pd.notna(row_atual["INICIO"]) else date.today()
                    dt_prz = row_atual["PRAZO FINAL"] if pd.notna(row_atual["PRAZO FINAL"]) else date.today()
                    
                    edit_inicio = col_ed1.date_input("📅 Data Início", value=dt_inc, format="DD/MM/YYYY")
                    edit_prazo = col_ed2.date_input("⏳ Prazo Final", value=dt_prz, format="DD/MM/YYYY")
                    
                    # Permite marcar ou limpar data de conclusão
                    tem_conclusao = st.checkbox("Marcar como Concluído?", value=pd.notna(row_atual["DATA CONCLUSAO"]))
                    dt_conc = row_atual["DATA CONCLUSAO"] if pd.notna(row_atual["DATA CONCLUSAO"]) else date.today()
                    edit_conclusao = col_ed3.date_input("✅ Data Conclusão", value=dt_conc, format="DD/MM/YYYY", disabled=not tem_conclusao)

                    btn_salvar_edicao = st.form_submit_button("💾 Salvar Alterações")

                    if btn_salvar_edicao:
                        df_bruto.loc[idx, "PROJETO"] = edit_nome
                        df_bruto.loc[idx, "TIPO"] = edit_tipo
                        df_bruto.loc[idx, "RESPONSAVEL"] = edit_resp
                        df_bruto.loc[idx, "INICIO"] = edit_inicio
                        df_bruto.loc[idx, "PRAZO FINAL"] = edit_prazo
                        df_bruto.loc[idx, "DATA CONCLUSAO"] = edit_conclusao if tem_conclusao else None

                        if salvar_dados_gsheets(df_bruto):
                            st.success(f"✅ Projeto '{edit_nome}' atualizado com sucesso!")
                            st.rerun()
            else:
                st.info("Nenhum projeto cadastrado para editar.")

        # ABA 3: REMOÇÃO DE PROJETO
        with aba_remover:
            st.subheader("🗑️ Remover Projeto")
            if not df_bruto.empty:
                projeto_rem = st.selectbox("Selecione o Projeto para Excluir:", df_bruto["PROJETO"].dropna().unique(), key="select_remover")
                
                st.warning(f"⚠️ Tem certeza que deseja excluir o projeto **'{projeto_rem}'**? Esta ação alterará a planilha no Google Sheets.")
                if st.button("❌ Confirmar Exclusão"):
                    df_novo = df_bruto[df_bruto["PROJETO"] != projeto_rem].copy()
                    if salvar_dados_gsheets(df_novo):
                        st.success(f"🗑️ Projeto '{projeto_rem}' removido com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhum projeto cadastrado para remover.")

    # ---------------------------------------------------------
    # 3. MÓDULO DE RELATÓRIOS
    # ---------------------------------------------------------
    elif opcao == "📋 Relatórios Ambientais":
        st.title("📋 Relatório Geral de Projetos e Serviços")

        st.subheader("🔍 Filtros de Consulta")
        col_f1, col_f2 = st.columns(2)

        responsaveis = ["Todos"] + list(df_calculado['RESPONSAVEL'].dropna().unique())
        resp_sel = col_f1.selectbox("👤 Responsável:", responsaveis)

        status_opcoes = ["Todos", "No Prazo", "Próximo do Vencimento", "Em Atraso", "Concluído"]
        status_sel = col_f2.selectbox("📌 Status do Projeto:", status_opcoes)

        df_filtrado = df_calculado.copy()
        if resp_sel != "Todos":
            df_filtrado = df_filtrado[df_filtrado['RESPONSAVEL'] == resp_sel]
        if status_sel != "Todos":
            df_filtrado = df_filtrado[df_filtrado['STATUS'] == status_sel]

        st.markdown("---")
        st.write(f"**Registros Localizados:** {len(df_filtrado)}")
        
        cols_exibicao = ["PROJETO", "TIPO", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO", "STATUS", "DIAS DE ATRASO"]
        df_filtrado_fmt = formatar_datas_para_exibicao(df_filtrado)
        
        st.dataframe(
            df_filtrado_fmt[cols_exibicao].style.apply(estilar_linha_inteira, axis=1),
            use_container_width=True
        )

        csv = df_filtrado_fmt[cols_exibicao].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Exportar Relatório em CSV",
            data=csv,
            file_name="relatorio_projetos_sanear.csv",
            mime="text/csv"
        )
