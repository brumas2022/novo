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

# Função para conectar e carregar os dados do Google Sheets usando st.connection
def carregar_dados_gsheets():
    try:
        # Inicializa a conexão declarada em [connections.gsheets] no secrets.toml
        conn = st.connection("gsheets", type=GSheetsConnection)
        
        # Lê a aba NovoPac da planilha PROJETOS_SANEAR
        df_projetos = conn.read(worksheet="NovoPac", ttl="0")
        
        # Lê a aba Servidor da planilha PROJETOS_SANEAR
        df_servidor = conn.read(worksheet="Servidor", ttl="0")

        # Limpeza básica nos nomes de colunas
        df_projetos.columns = df_projetos.columns.astype(str).str.strip()
        df_servidor.columns = df_servidor.columns.astype(str).str.strip()

        # Garante que a coluna de associação/responsável exista
        # Caso a aba Servidor tenha o mapeamento (ex: ID_PROJETO ou PROJETO para RESPONSAVEL)
        if "RESPONSAVEL" in df_servidor.columns and "RESPONSAVEL" not in df_projetos.columns:
            # Tenta realizar o merge por PROJETO se existir em ambas, ou traz diretamente se houver correspondência
            if "PROJETO" in df_servidor.columns and "PROJETO" in df_projetos.columns:
                df_projetos = df_projetos.merge(df_servidor[["PROJETO", "RESPONSAVEL"]], on="PROJETO", how="left")
            else:
                # Caso a aba Servidor seja apenas uma lista de responsáveis associada pelo índice
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

        return df_projetos[cols_obrigatorias]

    except Exception as e:
        st.error(f"⚠️ Erro ao conectar ao Google Sheets via secrets: {e}")
        return pd.DataFrame(columns=["PROJETO", "TIPO", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO"])

# Função para calcular o status e regras de vencimento
def calcular_status(df):
    hoje = date.today()
    status_list = []
    dias_atraso_list = []
    cor_linha_list = []

    for index, row in df.iterrows():
        prazo = row['PRAZO FINAL']
        conclusao = row['DATA CONCLUSAO']

        if pd.notna(conclusao):
            status_list.append("Concluído")
            dias_atraso_list.append(0)
            cor_linha_list.append("verde")
        elif pd.notna(prazo):
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
        ["📊 Dashboard & Prazos", "📋 Relatórios Ambientais"]
    )
    
    if st.sidebar.button("🔒 Sair / Logout"):
        st.session_state["autenticado"] = False
        st.rerun()

    # Carrega os dados diretamente do Google Sheets
    df_bruto = carregar_dados_gsheets()
    df_calculado = calcular_status(df_bruto)

    # 1. DASHBOARD & PRAZOS
    if opcao == "📊 Dashboard & Prazos":
        st.title("💧 Painel de Gestão de Projetos - SANEAR")
        
        # Métricas calculadas
        total_projetos = len(df_calculado)
        no_prazo = len(df_calculado[df_calculado['COR_LINHA'] == "verde"])
        prox_vencimento = len(df_calculado[df_calculado['COR_LINHA'] == "amarelo"])
        em_atraso = len(df_calculado[df_calculado['COR_LINHA'] == "vermelho"])

        # Quadrados / Cards customizados com as cores solicitadas
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
                    <div class="card-title">✅ No Prazo / Concluído</div>
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

        # Alerta de Projetos em Atraso
        df_atrasados = df_calculado[df_calculado['COR_LINHA'] == "vermelho"].sort_values(by="DIAS DE ATRASO", ascending=False)
        st.subheader("🚨 Projetos Com Prazos Vencidos (Linha Vermelha)")
        
        cols_exibicao = ["PROJETO", "TIPO", "RESPONSAVEL", "INICIO", "PRAZO FINAL", "DATA CONCLUSAO", "STATUS", "DIAS DE ATRASO"]
        
        if not df_atrasados.empty:
            df_atrasados_fmt = formatar_datas_para_exibicao(df_atrasados)
            st.dataframe(
                df_atrasados_fmt[cols_exibicao].style.apply(estilar_linha_inteira, axis=1),
                use_container_width=True
            )
        else:
            st.success("🌿 Excelente! Nenhum projeto em atraso no momento.")

        st.markdown("---")

        # Tabela Geral com coloração inteira por linha
        st.subheader("📋 Acompanhamento Geral de Prazos")
        df_geral_fmt = formatar_datas_para_exibicao(df_calculado)
        
        st.dataframe(
            df_geral_fmt[cols_exibicao].style.apply(estilar_linha_inteira, axis=1),
            use_container_width=True
        )

    # 2. MÓDULO DE RELATÓRIOS
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
