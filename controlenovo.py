import streamlit as st
import pandas as pd
import sqlite3
import datetime
import os

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA E ESTILO CSS (SISTEMA SIGA/MT - TOM VERDE CLARO)
# ==============================================================================
st.set_page_config(
    page_title="SIGA-MT | Gestão de Projetos e Licenciamento",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS personalizada com a paleta do SIGA/MT
st.markdown("""
<style>
    /* Fundo da aplicação em Verde Claro Suave */
    .stApp {
        background-color: #f4f8f5;
    }
    
    /* Cabeçalho no Estilo Governamental SIGA/MT */
    .siga-header {
        background: linear-gradient(135deg, #1b5e20 0%, #2e7d32 100%);
        color: white;
        padding: 20px 25px;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        margin-bottom: 25px;
    }
    .siga-header h1 {
        margin: 0;
        font-size: 1.8rem;
        font-weight: 700;
        color: #ffffff;
    }
    .siga-header p {
        margin: 5px 0 0 0;
        font-size: 0.95rem;
        opacity: 0.9;
    }

    /* Cards de Métricas Estilizados em Verde */
    .metric-card {
        background-color: #ffffff;
        border-left: 5px solid #2e7d32;
        border-radius: 8px;
        padding: 15px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        margin-bottom: 15px;
    }
    .metric-card .title {
        font-size: 0.85rem;
        color: #555555;
        font-weight: 600;
        text-transform: uppercase;
    }
    .metric-card .value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #1b5e20;
        margin-top: 5px;
    }

    /* Botões em Verde Floresta com Efeito Hover */
    .stButton>button {
        background-color: #2e7d32 !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 0.5rem 1.2rem !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        background-color: #1b5e20 !important;
        box-shadow: 0 4px 10px rgba(27, 94, 32, 0.3) !important;
    }

    /* Estilo de Abas (Tabs) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #e8f5e9;
        padding: 6px;
        border-radius: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        border-radius: 6px;
        color: #1b5e20;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2e7d32 !important;
        color: white !important;
    }

    /* Badges de Status */
    .badge-vencida {
        background-color: #ffebee;
        color: #c62828;
        padding: 4px 10px;
        border-radius: 12px;
        font-weight: bold;
        font-size: 0.8rem;
    }
    .badge-valida {
        background-color: #e8f5e9;
        color: #2e7d32;
        padding: 4px 10px;
        border-radius: 12px;
        font-weight: bold;
        font-size: 0.8rem;
    }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# 2. BANCO DE DADOS SQLITE & CARREGAMENTO INICIAL DA PLANILHA EXCEL
# ==============================================================================
DB_NAME = "projetos_sigamt.db"
EXCEL_FILE = "Protege-BETA EMPREENDIMENTO ATUAL (2).xls"

def get_connection():
    return sqlite3.connect(DB_NAME, check_same_thread=False)

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    # Tabela de Empresas / Clientes
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS empresas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT UNIQUE NOT NULL,
            cnpj TEXT,
            telefone TEXT,
            email TEXT,
            responsavel TEXT,
            municipio TEXT,
            uf TEXT DEFAULT 'MT',
            observacoes TEXT,
            data_cadastro TEXT
        )
    ''')
    
    # Tabela de Projetos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS projetos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            num_entrada TEXT,
            empresa_id INTEGER,
            cliente_nome TEXT,
            atividade TEXT,
            uf TEXT,
            municipio TEXT,
            num_processo TEXT,
            lp TEXT,
            li TEXT,
            lo TEXT,
            data_emissao TEXT,
            vencimento_licenca TEXT,
            vencimento_relatorio TEXT,
            prazo_licenca TEXT,
            prazo_relatorio TEXT,
            comentarios TEXT,
            valor_projeto REAL DEFAULT 0.0,
            valor_pago REAL DEFAULT 0.0,
            aditivos REAL DEFAULT 0.0,
            descontos REAL DEFAULT 0.0,
            valor_saldo REAL DEFAULT 0.0,
            status_pagamento TEXT DEFAULT 'Em Aberto',
            status_licenca TEXT,
            data_cadastro TEXT,
            FOREIGN KEY(empresa_id) REFERENCES empresas(id) ON DELETE CASCADE
        )
    ''')
    conn.commit()
    
    # Carga inicial se o banco estiver vazio
    cursor.execute("SELECT COUNT(*) FROM empresas")
    if cursor.fetchone()[0] == 0 and os.path.exists(EXCEL_FILE):
        cargar_dados_iniciais_excel(conn)
        
    conn.close()

def cargar_dados_iniciais_excel(conn):
    try:
        cursor = conn.cursor()
        df_cad_raw = pd.read_excel(EXCEL_FILE, sheet_name='CADASTRO', header=2)
        col_names = df_cad_raw.iloc[0].values
        df_clean = df_cad_raw.iloc[1:].copy()
        df_clean.columns = [str(c).strip() for c in col_names]
        df_clean = df_clean.dropna(subset=['CLIENTE'])

        today_str = datetime.date.today().strftime('%Y-%m-%d')

        # Insere Empresas
        unique_clients = df_clean['CLIENTE'].unique()
        for client in unique_clients:
            client_str = str(client).strip()
            cursor.execute('''
                INSERT OR IGNORE INTO empresas (nome, uf, data_cadastro)
                VALUES (?, 'MT', ?)
            ''', (client_str, today_str))
        conn.commit()

        # Mapeia IDs
        cursor.execute("SELECT id, nome FROM empresas")
        emp_map = {row[1]: row[0] for row in cursor.fetchall()}

        # Insere Projetos
        for _, row in df_clean.iterrows():
            cli_name = str(row['CLIENTE']).strip()
            emp_id = emp_map.get(cli_name)

            dt_emissao = pd.to_datetime(row['DATA DE EMISSÃO'], errors='coerce')
            dt_emissao_str = dt_emissao.strftime('%Y-%m-%d') if pd.notna(dt_emissao) else None

            dt_venc = pd.to_datetime(row['VENCIMENTO LICENÇA'], errors='coerce')
            dt_venc_str = dt_venc.strftime('%Y-%m-%d') if pd.notna(dt_venc) else None

            dt_venc_rel = pd.to_datetime(row['VENCIMENTO RELATÓRIO'], errors='coerce')
            dt_venc_rel_str = dt_venc_rel.strftime('%Y-%m-%d') if pd.notna(dt_venc_rel) else None

            status_lic = str(row['PRAZO DA LICENÇA']) if pd.notna(row['PRAZO DA LICENÇA']) else 'Válida'

            # Valores financeiros padrão para demonstração inicial
            val_proj = 50000.00
            val_pago = 25000.00
            val_saldo = val_proj - val_pago

            cursor.execute('''
                INSERT INTO projetos (
                    num_entrada, empresa_id, cliente_nome, atividade, uf, municipio,
                    num_processo, lp, li, lo, data_emissao, vencimento_licenca,
                    vencimento_relatorio, prazo_licenca, prazo_relatorio, comentarios,
                    valor_projeto, valor_pago, aditivos, descontos, valor_saldo,
                    status_pagamento, status_licenca, data_cadastro
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0.0, 0.0, ?, 'Parcialmente Pago', ?, ?)
            ''', (
                str(row['Nº DE ENTRADA']) if pd.notna(row['Nº DE ENTRADA']) else None,
                emp_id, cli_name,
                str(row['ATIVIDADE']) if pd.notna(row['ATIVIDADE']) else '',
                str(row['UF']) if pd.notna(row['UF']) else 'MT',
                str(row['MUNICÍPIO']) if pd.notna(row['MUNICÍPIO']) else '',
                str(row['NÚMERO DO PROCESSO']) if pd.notna(row['NÚMERO DO PROCESSO']) else '',
                'X' if str(row['LP']).strip().upper() == 'X' else '',
                'X' if str(row['LI']).strip().upper() == 'X' else '',
                'X' if str(row['LO']).strip().upper() == 'X' else '',
                dt_emissao_str, dt_venc_str, dt_venc_rel_str, status_lic,
                str(row['PRAZO DO RELATÓRIO']) if pd.notna(row['PRAZO DO RELATÓRIO']) else '',
                str(row['COMENTÁRIOS']) if pd.notna(row['COMENTÁRIOS']) else '',
                val_proj, val_pago, val_saldo, status_lic, today_str
            ))
        conn.commit()
    except Exception as e:
        st.error(f"Erro ao carregar planilha inicial: {e}")

# Executa inicialização
init_db()


# ==============================================================================
# 3. LISTA DE MUNICÍPIOS DE MATO GROSSO (SISTEMA SIGA/MT)
# ==============================================================================
MUNICIPIOS_MT = [
    "Acorizal", "Água Boa", "Alta Floresta", "Alto Araguaia", "Alto Boa Vista", 
    "Alto Garças", "Alto Paraguai", "Alto Taquari", "Apiacás", "Araguaiana", 
    "Aragwaita", "Araputanga", "Arenápolis", "Aripuanã", "Barão de Melgaço", 
    "Barra do Bugres", "Barra do Garças", "Bom Jesus do Araguaia", "Brasnorte", 
    "Cáceres", "Campinápolis", "Campo Novo do Parecis", "Campo Verde", "Campos de Júlio", 
    "Canabrava do Norte", "Canarana", "Carlinda", "Castanheira", "Chapada dos Guimarães", 
    "Cláudia", "Cocalinho", "Colíder", "Colniza", "Comodoro", "Confresa", 
    "Conquista D'Oeste", "Cotriguaçu", "Cuiabá", "Curvelândia", "Denise", 
    "Diamantino", "Dom Aquino", "Feliz Natal", "Figueirópolis D'Oeste", "Gaúcha do Norte", 
    "General Carneiro", "Glória D'Oeste", "Guarantã do Norte", "Guiratinga", "Indiavaí", 
    "Ipiranga do Norte", "Itanhangá", "Itaúba", "Itiquira", "Jaciara", "Jangada", 
    "Jauru", "Juara", "Juína", "Juruena", "Juscimeira", "Lucas do Rio Verde", 
    "Luciara", "Marcelândia", "Matupá", "Mirassol d'Oeste", "Nobres", "Nortelândia", 
    "Nossa Senhora do Livramento", "Nova Bandeirantes", "Nova Brasilândia", 
    "Nova Canaã do Norte", "Nova Guarita", "Nova Lacerda", "Nova Marilândia", 
    "Nova Maringá", "Nova Monte Verde", "Nova Mutum", "Nova Nazaré", "Nova Olímpia", 
    "Nova Santa Helena", "Nova Ubiratã", "Nova Xavantina", "Novo Horizonte do Norte", 
    "Novo Mundo", "Novo Santo Antônio", "Novo São Joaquim", "Paranaíta", "Paranatinga", 
    "Pedra Preta", "Peixoto de Azevedo", "Planalto da Serra", "Poconé", "Pontal do Araguaia", 
    "Ponte e Lacerda", "Pontes e Lacerda", "Porto Alegre do Norte", "Porto Esperidião", 
    "Porto Estrela", "Porto dos Gaúchos", "Poxoréu", "Primavera do Leste", "Querência", 
    "Reserva do Cabaçal", "Ribeirão Cascalheira", "Ribeirãozinho", "Rio Branco", "Rondolândia", 
    "Rondonópolis", "Rosário Oeste", "Salto do Céu", "Santa Carmem", "Santa Cruz do Xingu", 
    "Santa Luciene", "Santa Rita do Trivelato", "Santa Terezinha", "Santo Antônio do Leste", 
    "Santo Antônio do Leverger", "São Félix do Araguaia", "São José do Povo", 
    "São José do Rio Claro", "São José do Xingu", "São José dos Quatro Marcos", 
    "São Pedro da Cipa", "Sapezal", "Sinop", "Sorriso", "Tabaporã", "Tangará da Serra", 
    "Tapas", "Terra Nova do Norte", "Tesouro", "Torixoréu", "União do Sul", 
    "Vale de São Domingos", "Várzea Grande", "Vera", "Vila Bela da Santíssima Trindade", "Vila Rica"
]


# ==============================================================================
# 4. FUNÇÕES DE BANCO DE DADOS (CRUD EMPRESAS E PROJETOS)
# ==============================================================================

# --- EMPRESAS ---
def list_empresas():
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM empresas ORDER BY nome ASC", conn)
    conn.close()
    return df

def add_empresa(nome, cnpj, telefone, email, responsavel, municipio, uf, observacoes):
    conn = get_connection()
    cursor = conn.cursor()
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    try:
        cursor.execute('''
            INSERT INTO empresas (nome, cnpj, telefone, email, responsavel, municipio, uf, observacoes, data_cadastro)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (nome, cnpj, telefone, email, responsavel, municipio, uf, observacoes, today_str))
        conn.commit()
        conn.close()
        return True, "Empresa cadastrada com sucesso!"
    except sqlite3.IntegrityError:
        conn.close()
        return False, "Já existe uma empresa cadastrada com este nome."
    except Exception as e:
        conn.close()
        return False, f"Erro ao cadastrar empresa: {e}"

def update_empresa(empresa_id, nome, cnpj, telefone, email, responsavel, municipio, uf, observacoes):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute('''
            UPDATE empresas 
            SET nome=?, cnpj=?, telefone=?, email=?, responsavel=?, municipio=?, uf=?, observacoes=?
            WHERE id=?
        ''', (nome, cnpj, telefone, email, responsavel, municipio, uf, observacoes, empresa_id))
        
        # Atualiza o nome do cliente na tabela de projetos
        cursor.execute("UPDATE projetos SET cliente_nome=? WHERE empresa_id=?", (nome, empresa_id))
        
        conn.commit()
        conn.close()
        return True, "Dados da empresa atualizados com sucesso!"
    except Exception as e:
        conn.close()
        return False, f"Erro ao atualizar empresa: {e}"

def delete_empresa(empresa_id):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM projetos WHERE empresa_id=?", (empresa_id,))
        cursor.execute("DELETE FROM empresas WHERE id=?", (empresa_id,))
        conn.commit()
        conn.close()
        return True, "Empresa e projetos associados removidos com sucesso!"
    except Exception as e:
        conn.close()
        return False, f"Erro ao remover empresa: {e}"


# --- PROJETOS ---
def list_projetos(empresa_id=None, busca=None):
    conn = get_connection()
    query = """
        SELECT p.*, e.cnpj as empresa_cnpj, e.email as empresa_email 
        FROM projetos p
        LEFT JOIN empresas e ON p.empresa_id = e.id
        WHERE 1=1
    """
    params = []
    
    if empresa_id:
        query += " AND p.empresa_id = ?"
        params.append(empresa_id)
        
    if busca:
        query += " AND (p.cliente_nome LIKE ? OR p.atividade LIKE ? OR p.num_processo LIKE ? OR p.municipio LIKE ?)"
        term = f"%{busca}%"
        params.extend([term, term, term, term])
        
    query += " ORDER BY p.id DESC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def add_projeto(num_entrada, empresa_id, atividade, uf, municipio, num_processo, 
                lp, li, lo, data_emissao, vencimento_licenca, vencimento_relatorio, 
                prazo_licenca, prazo_relatorio, comentarios, valor_projeto, valor_pago, aditivos, descontos, status_pagamento):
    conn = get_connection()
    cursor = conn.cursor()
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    
    # Busca nome do cliente
    cursor.execute("SELECT nome FROM empresas WHERE id=?", (empresa_id,))
    emp_res = cursor.fetchone()
    cliente_nome = emp_res[0] if emp_res else "Empresa Desconhecida"
    
    # Cálculo automático do saldo
    valor_saldo = float(valor_projeto) + float(aditivos) - float(descontos) - float(valor_pago)
    
    try:
        cursor.execute('''
            INSERT INTO projetos (
                num_entrada, empresa_id, cliente_nome, atividade, uf, municipio,
                num_processo, lp, li, lo, data_emissao, vencimento_licenca,
                vencimento_relatorio, prazo_licenca, prazo_relatorio, comentarios,
                valor_projeto, valor_pago, aditivos, descontos, valor_saldo,
                status_pagamento, status_licenca, data_cadastro
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            num_entrada, empresa_id, cliente_nome, atividade, uf, municipio,
            num_processo, lp, li, lo, data_emissao, vencimento_licenca,
            vencimento_relatorio, prazo_licenca, prazo_relatorio, comentarios,
            valor_projeto, valor_pago, aditivos, descontos, valor_saldo,
            status_pagamento, prazo_licenca, today_str
        ))
        conn.commit()
        conn.close()
        return True, "Projeto cadastrado com sucesso!"
    except Exception as e:
        conn.close()
        return False, f"Erro ao cadastrar projeto: {e}"

def update_projeto(projeto_id, num_entrada, empresa_id, atividade, uf, municipio, num_processo, 
                   lp, li, lo, data_emissao, vencimento_licenca, vencimento_relatorio, 
                   prazo_licenca, prazo_relatorio, comentarios, valor_projeto, valor_pago, aditivos, descontos, status_pagamento):
    conn = get_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT nome FROM empresas WHERE id=?", (empresa_id,))
    emp_res = cursor.fetchone()
    cliente_nome = emp_res[0] if emp_res else "Empresa Desconhecida"
    
    valor_saldo = float(valor_projeto) + float(aditivos) - float(descontos) - float(valor_pago)
    
    try:
        cursor.execute('''
            UPDATE projetos SET
                num_entrada=?, empresa_id=?, cliente_nome=?, atividade=?, uf=?, municipio=?,
                num_processo=?, lp=?, li=?, lo=?, data_emissao=?, vencimento_licenca=?,
                vencimento_relatorio=?, prazo_licenca=?, prazo_relatorio=?, comentarios=?,
                valor_projeto=?, valor_pago=?, aditivos=?, descontos=?, valor_saldo=?,
                status_pagamento=?, status_licenca=?
            WHERE id=?
        ''', (
            num_entrada, empresa_id, cliente_nome, atividade, uf, municipio,
            num_processo, lp, li, lo, data_emissao, vencimento_licenca,
            vencimento_relatorio, prazo_licenca, prazo_relatorio, comentarios,
            valor_projeto, valor_pago, aditivos, descontos, valor_saldo,
            status_pagamento, prazo_licenca, projeto_id
        ))
        conn.commit()
        conn.close()
        return True, "Projeto atualizado com sucesso!"
    except Exception as e:
        conn.close()
        return False, f"Erro ao atualizar projeto: {e}"

def delete_projeto(projeto_id):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM projetos WHERE id=?", (projeto_id,))
        conn.commit()
        conn.close()
        return True, "Projeto removido com sucesso!"
    except Exception as e:
        conn.close()
        return False, f"Erro ao remover projeto: {e}"


# ==============================================================================
# 5. ESTRUTURA DO MENU SIDEBAR E NAVEGAÇÃO
# ==============================================================================
st.sidebar.image("https://img.icons8.com/color/96/000000/tree-structure.png", width=70)
st.sidebar.title("SIGA-MT | PROTEGE")
st.sidebar.caption("Sistema de Gestão de Projetos e Licenciamento Ambiental")
st.sidebar.markdown("---")

modulo = st.sidebar.radio(
    "📌 Selecione o Módulo:",
    [
        "📊 Dashboard Geral",
        "🏢 Módulo Empresas",
        "📁 Módulo Projetos",
        "📄 Relatórios & Exportação",
        "⚙️ Configurações"
    ]
)

st.sidebar.markdown("---")
st.sidebar.info("💡 **Estado de Mato Grosso**\n\nGoverno do Estado de Mato Grosso\nGestão de Licenciamento Ambiental")


# ==============================================================================
# MÓDULO 1: DASHBOARD GERAL
# ==============================================================================
if modulo == "📊 Dashboard Geral":
    st.markdown("""
        <div class="siga-header">
            <h1>🌿 SIGA-MT — Painel Geral de Projetos e Licenciamento</h1>
            <p>Visão consolidada do controle de andamento, licenças ambientais e financeiros dos empreendimentos.</p>
        </div>
    """, unsafe_allow_html=True)
    
    df_p = list_projetos()
    df_e = list_empresas()
    
    # Métricas Principais em Cards Estilizados
    tot_proj = len(df_p)
    tot_emp = len(df_e)
    val_total = df_p['valor_projeto'].sum()
    val_pago = df_p['valor_pago'].sum()
    val_saldo = df_p['valor_saldo'].sum()
    lic_vencidas = len(df_p[df_p['prazo_licenca'].str.contains("Vencida", case=False, na=False)])

    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="title">Total Projetos</div>
                <div class="value">{tot_proj}</div>
            </div>
        """, unsafe_allow_html=True)
        
    with col2:
        st.markdown(f"""
            <div class="metric-card">
                <div class="title">Empresas Cadastradas</div>
                <div class="value">{tot_emp}</div>
            </div>
        """, unsafe_allow_html=True)
        
    with col3:
        st.markdown(f"""
            <div class="metric-card">
                <div class="title">Valor Total (R$)</div>
                <div class="value">R$ {val_total:,.2f}</div>
            </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
            <div class="metric-card">
                <div class="title">Saldo Pendente (R$)</div>
                <div class="value">R$ {val_saldo:,.2f}</div>
            </div>
        """, unsafe_allow_html=True)

    with col5:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #c62828;">
                <div class="title">Licenças Vencidas</div>
                <div class="value" style="color: #c62828;">{lic_vencidas}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("### 📈 Resumo Gráfico e Indicadores")
    g_col1, g_col2 = st.columns(2)
    
    with g_col1:
        st.markdown("#### Projetos por Município (MT)")
        if not df_p.empty and 'municipio' in df_p.columns:
            mun_counts = df_p['municipio'].value_counts().head(10)
            st.bar_chart(mun_counts)
            
    with g_col2:
        st.markdown("#### Status de Pagamento dos Projetos")
        if not df_p.empty and 'status_pagamento' in df_p.columns:
            pag_counts = df_p['status_pagamento'].value_counts()
            st.bar_chart(pag_counts)

    st.markdown("### 📋 Visão Rápida dos Últimos Projetos Cadastrados")
    st.dataframe(
        df_p[['id', 'cliente_nome', 'atividade', 'municipio', 'num_processo', 'prazo_licenca', 'valor_projeto', 'valor_saldo']].head(10),
        use_container_width=True,
        hide_index=True
    )


# ==============================================================================
# MÓDULO 2: GESTÃO DE EMPRESAS (CRUD COMPLETO)
# ==============================================================================
elif modulo == "🏢 Módulo Empresas":
    st.markdown("""
        <div class="siga-header">
            <h1>🏢 Gestão de Empresas e Clientes</h1>
            <p>Cadastro, consulta, edição e remoção das empresas empreendedoras no SIGA-MT.</p>
        </div>
    """, unsafe_allow_html=True)
    
    tab_list_emp, tab_add_emp, tab_edit_emp, tab_del_emp = st.tabs([
        "📋 Consultar Empresas", 
        "➕ Cadastrar Nova Empresa", 
        "✏️ Editar Empresa", 
        "🗑️ Remover Empresa"
    ])
    
    # --- 1. LISTAR EMPRESAS ---
    with tab_list_emp:
        st.markdown("### Lista de Empresas Cadastradas")
        busca_emp = st.text_input("🔍 Buscar por Nome, CNPJ ou Município:", key="busca_emp_input")
        
        df_e = list_empresas()
        if busca_emp:
            df_e = df_e[
                df_e['nome'].str.contains(busca_emp, case=False, na=False) |
                df_e['cnpj'].str.contains(busca_emp, case=False, na=False) |
                df_e['municipio'].str.contains(busca_emp, case=False, na=False)
            ]
            
        st.dataframe(df_e, use_container_width=True, hide_index=True)
        
        st.markdown("#### 📁 Detalhes da Empresa e Projetos Associados")
        if not df_e.empty:
            emp_sel_id = st.selectbox("Selecione uma empresa para visualizar seus projetos:", df_e['id'].tolist(), format_func=lambda x: df_e[df_e['id']==x]['nome'].values[0])
            if emp_sel_id:
                df_projs_emp = list_projetos(empresa_id=emp_sel_id)
                st.write(f"**Projetos vinculados ({len(df_projs_emp)}):**")
                st.dataframe(df_projs_emp[['id', 'num_processo', 'atividade', 'municipio', 'prazo_licenca', 'valor_projeto', 'valor_saldo']], use_container_width=True, hide_index=True)

    # --- 2. CADASTRAR EMPRESA ---
    with tab_add_emp:
        st.markdown("### ➕ Formulario de Cadastro de Nova Empresa")
        with st.form("form_add_empresa", clear_on_submit=True):
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                nome_emp = st.text_input("Nome Razão Social / Cliente *")
                cnpj_emp = st.text_input("CNPJ / CPF")
                telefone_emp = st.text_input("Telefone de Contato")
                email_emp = st.text_input("E-mail de Contato")
            with col_e2:
                resp_emp = st.text_input("Responsável Técnico / Contato")
                mun_emp = st.selectbox("Município Sede", [""] + MUNICIPIOS_MT)
                uf_emp = st.text_input("UF", value="MT")
                obs_emp = st.text_area("Observações Adicionais", height=68)
                
            submitted = st.form_submit_button("💾 Salvar Empresa")
            if submitted:
                if not nome_emp.strip():
                    st.error("O campo Nome Razão Social é obrigatório.")
                else:
                    success, msg = add_empresa(nome_emp.strip(), cnpj_emp, telefone_emp, email_emp, resp_emp, mun_emp, uf_emp, obs_emp)
                    if success:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

    # --- 3. EDITAR EMPRESA ---
    with tab_edit_emp:
        st.markdown("### ✏️ Editar Dados da Empresa")
        df_e_all = list_empresas()
        if df_e_all.empty:
            st.warning("Nenhuma empresa cadastrada para editar.")
        else:
            emp_edit_id = st.selectbox("Selecione a Empresa para Editar:", df_e_all['id'].tolist(), format_func=lambda x: df_e_all[df_e_all['id']==x]['nome'].values[0])
            row_e = df_e_all[df_e_all['id'] == emp_edit_id].iloc[0]
            
            with st.form("form_edit_empresa"):
                col_ee1, col_ee2 = st.columns(2)
                with col_ee1:
                    nome_ee = st.text_input("Nome Razão Social / Cliente *", value=row_e['nome'])
                    cnpj_ee = st.text_input("CNPJ / CPF", value=row_e['cnpj'] or "")
                    telefone_ee = st.text_input("Telefone de Contato", value=row_e['telefone'] or "")
                    email_ee = st.text_input("E-mail de Contato", value=row_e['email'] or "")
                with col_ee2:
                    resp_ee = st.text_input("Responsável Técnico / Contato", value=row_e['responsavel'] or "")
                    
                    mun_idx = MUNICIPIOS_MT.index(row_e['municipio']) if row_e['municipio'] in MUNICIPIOS_MT else 0
                    mun_ee = st.selectbox("Município Sede", MUNICIPIOS_MT, index=mun_idx)
                    uf_ee = st.text_input("UF", value=row_e['uf'] or "MT")
                    obs_ee = st.text_area("Observações Adicionais", value=row_e['observacoes'] or "", height=68)
                    
                sub_edit = st.form_submit_button("🔄 Atualizar Dados da Empresa")
                if sub_edit:
                    if not nome_ee.strip():
                        st.error("O campo Nome da Empresa não pode ficar vazio.")
                    else:
                        success, msg = update_empresa(emp_edit_id, nome_ee.strip(), cnpj_ee, telefone_ee, email_ee, resp_ee, mun_ee, uf_ee, obs_ee)
                        if success:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)

    # --- 4. REMOVER EMPRESA ---
    with tab_del_emp:
        st.markdown("### 🗑️ Remover Empresa")
        df_e_del = list_empresas()
        if df_e_del.empty:
            st.info("Nenhuma empresa disponível para remoção.")
        else:
            emp_del_id = st.selectbox("Selecione a Empresa a ser Removida:", df_e_del['id'].tolist(), format_func=lambda x: df_e_del[df_e_del['id']==x]['nome'].values[0])
            
            df_linked = list_projetos(empresa_id=emp_del_id)
            st.warning(f"⚠️ **Atenção:** Esta ação excluirá a empresa e **{len(df_linked)} projeto(s)** vinculados a ela de forma permanente.")
            
            if st.button("❌ Confirmar Exclusão da Empresa"):
                success, msg = delete_empresa(emp_del_id)
                if success:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)


# ==============================================================================
# MÓDULO 3: GESTÃO DE PROJETOS (CRUD + INSERÇÃO DE VALORES FINANCEIROS)
# ==============================================================================
elif modulo == "📁 Módulo Projetos":
    st.markdown("""
        <div class="siga-header">
            <h1>📁 Gestão de Projetos e Controle Financeiro</h1>
            <p>Cadastro, acompanhamento de prazos de licenças ambientais e gestão de valores financeiros dos projetos.</p>
        </div>
    """, unsafe_allow_html=True)
    
    tab_list_p, tab_add_p, tab_edit_p, tab_del_p = st.tabs([
        "📋 Consultar Projetos", 
        "➕ Cadastrar Novo Projeto", 
        "✏️ Editar Projeto & Valores", 
        "🗑️ Remover Projeto"
    ])
    
    # --- 1. LISTAR PROJETOS ---
    with tab_list_p:
        st.markdown("### Lista e Consulta de Projetos")
        c_f1, c_f2 = st.columns([3, 1])
        with c_f1:
            busca_proj = st.text_input("🔍 Buscar por Cliente, Processo ou Atividade:")
        with c_f2:
            filtro_lic = st.selectbox("Filtrar por Licença:", ["Todos", "Licença Vencida", "Válida", "LP", "LI", "LO"])
            
        df_p_view = list_projetos(busca=busca_proj)
        
        if filtro_lic != "Todos":
            if filtro_lic in ["LP", "LI", "LO"]:
                col_check = filtro_lic.lower()
                df_p_view = df_p_view[df_p_view[col_check] == 'X']
            else:
                df_p_view = df_p_view[df_p_view['prazo_licenca'].str.contains(filtro_lic, case=False, na=False)]
                
        # Exibição Formatada de Valores
        if not df_p_view.empty:
            df_display = df_p_view.copy()
            df_display['Valor Projeto (R$)'] = df_display['valor_projeto'].apply(lambda x: f"R$ {x:,.2f}")
            df_display['Saldo (R$)'] = df_display['valor_saldo'].apply(lambda x: f"R$ {x:,.2f}")
            
            st.dataframe(
                df_display[[
                    'id', 'cliente_nome', 'atividade', 'municipio', 'num_processo', 
                    'lp', 'li', 'lo', 'vencimento_licenca', 'prazo_licenca', 
                    'Valor Projeto (R$)', 'Saldo (R$)', 'status_pagamento'
                ]],
                use_container_width=True,
                hide_index=True
            )
            
            st.markdown("#### 🔍 Detalhes Completos do Projeto Selecionado")
            p_sel_id = st.selectbox("Selecione o Projeto para ver os detalhes completos:", df_p_view['id'].tolist(), format_func=lambda x: f"ID #{x} - {df_p_view[df_p_view['id']==x]['cliente_nome'].values[0]} ({df_p_view[df_p_view['id']==x]['atividade'].values[0]})")
            
            if p_sel_id:
                p_row = df_p_view[df_p_view['id'] == p_sel_id].iloc[0]
                with st.expander("📄 Ver Ficha Técnica Completa do Projeto", expanded=True):
                    col_dt1, col_dt2, col_dt3 = st.columns(3)
                    with col_dt1:
                        st.write(f"**Cliente:** {p_row['cliente_nome']}")
                        st.write(f"**Atividade:** {p_row['atividade']}")
                        st.write(f"**Nº Processo:** {p_row['num_processo']}")
                        st.write(f"**Município / UF:** {p_row['municipio']} / {p_row['uf']}")
                    with col_dt2:
                        st.write(f"**Licenças:** LP: [{p_row['lp'] or ' '}] | LI: [{p_row['li'] or ' '}] | LO: [{p_row['lo'] or ' '}]")
                        st.write(f"**Data de Emissão:** {p_row['data_emissao'] or 'N/I'}")
                        st.write(f"**Vencimento Licença:** {p_row['vencimento_licenca'] or 'N/I'}")
                        st.write(f"**Status Licença:** {p_row['prazo_licenca']}")
                    with col_dt3:
                        st.markdown("**💰 Resumo Financeiro:**")
                        st.write(f"**Valor do Projeto:** R$ {p_row['valor_projeto']:,.2f}")
                        st.write(f"**Valor Pago:** R$ {p_row['valor_pago']:,.2f}")
                        st.write(f"**Aditivos:** R$ {p_row['aditivos']:,.2f} | **Descontos:** R$ {p_row['descontos']:,.2f}")
                        st.write(f"**Saldo Restante:** R$ {p_row['valor_saldo']:,.2f}")
                        st.write(f"**Status Pagamento:** {p_row['status_pagamento']}")
                    
                    if p_row['comentarios']:
                        st.info(f"**Comentários / Observações:** {p_row['comentarios']}")

    # --- 2. CADASTRAR PROJETO ---
    with tab_add_p:
        st.markdown("### ➕ Novo Cadastro de Projeto")
        df_empresas_all = list_empresas()
        
        if df_empresas_all.empty:
            st.warning("É necessário cadastrar ao menos uma Empresa/Cliente antes de criar projetos.")
        else:
            with st.form("form_add_projeto", clear_on_submit=True):
                st.markdown("#### 1. Identificação do Empreendimento")
                col_p1, col_p2, col_p3 = st.columns(3)
                
                with col_p1:
                    num_ent = st.text_input("Nº de Entrada")
                    emp_id_sel = st.selectbox("Empresa / Cliente *", df_empresas_all['id'].tolist(), format_func=lambda x: df_empresas_all[df_empresas_all['id']==x]['nome'].values[0])
                    atividade_p = st.text_input("Atividade / Empreendimento *")
                with col_p2:
                    num_proc = st.text_input("Nº do Processo")
                    mun_p = st.selectbox("Município do Projeto", MUNICIPIOS_MT)
                    uf_p = st.text_input("UF", value="MT")
                with col_p3:
                    st.markdown("**Tipos de Licença Possuídos:**")
                    lp_check = st.checkbox("LP - Licença Prévia")
                    li_check = st.checkbox("LI - Licença de Instalação")
                    lo_check = st.checkbox("LO - Licença de Operação")
                
                st.markdown("---")
                st.markdown("#### 2. Datas e Licenciamento Ambiental")
                col_d1, col_d2, col_d3 = st.columns(3)
                with col_d1:
                    dt_emissao_p = st.date_input("Data de Emissão", value=datetime.date.today())
                with col_d2:
                    dt_venc_p = st.date_input("Vencimento da Licença", value=datetime.date.today() + datetime.timedelta(days=365))
                with col_d3:
                    prazo_lic_p = st.selectbox("Status da Licença", ["Válida", "Licença Vencida", "Em Renovação", "Pendente"])
                
                coment_p = st.text_area("Comentários / Observações do Licenciamento", height=70)

                st.markdown("---")
                st.markdown("#### 3. Gestão Financeira e Valores (R$)")
                col_v1, col_v2, col_v3, col_v4 = st.columns(4)
                
                with col_v1:
                    val_proj_p = st.number_input("Valor Total do Projeto (R$)", min_value=0.0, value=10000.0, step=500.0)
                with col_v2:
                    val_pago_p = st.number_input("Valor Pago (R$)", min_value=0.0, value=0.0, step=500.0)
                with col_v3:
                    val_adit_p = st.number_input("Aditivos (R$)", min_value=0.0, value=0.0, step=100.0)
                    val_desc_p = st.number_input("Descontos (R$)", min_value=0.0, value=0.0, step=100.0)
                with col_v4:
                    status_pag_p = st.selectbox("Status do Pagamento", ["Em Aberto", "Parcialmente Pago", "Quitado", "Atrasado"])

                sub_add_p = st.form_submit_button("💾 Cadastrar Projeto")
                if sub_add_p:
                    if not atividade_p.strip():
                        st.error("O campo Atividade é obrigatório.")
                    else:
                        success, msg = add_projeto(
                            num_ent, emp_id_sel, atividade_p.strip(), uf_p, mun_p, num_proc,
                            'X' if lp_check else '', 'X' if li_check else '', 'X' if lo_check else '',
                            dt_emissao_p.strftime('%Y-%m-%d'), dt_venc_p.strftime('%Y-%m-%d'), None,
                            prazo_lic_p, None, coment_p,
                            val_proj_p, val_pago_p, val_adit_p, val_desc_p, status_pag_p
                        )
                        if success:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)

    # --- 3. EDITAR PROJETO & VALORES ---
    with tab_edit_p:
        st.markdown("### ✏️ Editar Dados e Valores do Projeto")
        df_projs_all = list_projetos()
        
        if df_projs_all.empty:
            st.warning("Nenhum projeto cadastrado para edição.")
        else:
            p_edit_id = st.selectbox("Selecione o Projeto para Editar:", df_projs_all['id'].tolist(), format_func=lambda x: f"ID #{x} - {df_projs_all[df_projs_all['id']==x]['cliente_nome'].values[0]} ({df_projs_all[df_projs_all['id']==x]['atividade'].values[0]})")
            p_row_e = df_projs_all[df_projs_all['id'] == p_edit_id].iloc[0]
            
            df_emp_list = list_empresas()
            emp_curr_id = p_row_e['empresa_id'] if p_row_e['empresa_id'] in df_emp_list['id'].values else df_emp_list['id'].iloc[0]
            
            with st.form("form_edit_projeto"):
                st.markdown("#### 1. Identificação")
                col_pe1, col_pe2, col_pe3 = st.columns(3)
                
                with col_pe1:
                    num_ent_e = st.text_input("Nº de Entrada", value=p_row_e['num_entrada'] or "")
                    emp_id_e = st.selectbox("Empresa / Cliente", df_emp_list['id'].tolist(), index=list(df_emp_list['id']).index(emp_curr_id), format_func=lambda x: df_emp_list[df_emp_list['id']==x]['nome'].values[0])
                    atividade_pe = st.text_input("Atividade / Empreendimento", value=p_row_e['atividade'] or "")
                with col_pe2:
                    num_proc_e = st.text_input("Nº do Processo", value=p_row_e['num_processo'] or "")
                    mun_p_idx = MUNICIPIOS_MT.index(p_row_e['municipio']) if p_row_e['municipio'] in MUNICIPIOS_MT else 0
                    mun_pe = st.selectbox("Município do Projeto", MUNICIPIOS_MT, index=mun_p_idx)
                    uf_pe = st.text_input("UF", value=p_row_e['uf'] or "MT")
                with col_pe3:
                    st.markdown("**Tipos de Licença:**")
                    lp_check_e = st.checkbox("LP", value=(p_row_e['lp'] == 'X'))
                    li_check_e = st.checkbox("LI", value=(p_row_e['li'] == 'X'))
                    lo_check_e = st.checkbox("LO", value=(p_row_e['lo'] == 'X'))

                st.markdown("---")
                st.markdown("#### 2. Datas e Licenciamento")
                col_de1, col_de2, col_de3 = st.columns(3)
                
                dt_em_val = pd.to_datetime(p_row_e['data_emissao']).date() if pd.notna(p_row_e['data_emissao']) else datetime.date.today()
                dt_vc_val = pd.to_datetime(p_row_e['vencimento_licenca']).date() if pd.notna(p_row_e['vencimento_licenca']) else datetime.date.today()
                
                with col_de1:
                    dt_emissao_pe = st.date_input("Data de Emissão", value=dt_em_val)
                with col_de2:
                    dt_venc_pe = st.date_input("Vencimento da Licença", value=dt_vc_val)
                with col_de3:
                    prazo_opts = ["Válida", "Licença Vencida", "Em Renovação", "Pendente"]
                    pr_idx = prazo_opts.index(p_row_e['prazo_licenca']) if p_row_e['prazo_licenca'] in prazo_opts else 0
                    prazo_lic_pe = st.selectbox("Status da Licença", prazo_opts, index=pr_idx)
                
                coment_pe = st.text_area("Comentários", value=p_row_e['comentarios'] or "", height=70)

                st.markdown("---")
                st.markdown("#### 3. Valores Financeiros (R$)")
                col_ve1, col_ve2, col_ve3, col_ve4 = st.columns(4)
                
                with col_ve1:
                    val_proj_pe = st.number_input("Valor Total do Projeto (R$)", min_value=0.0, value=float(p_row_e['valor_projeto'] or 0.0), step=500.0)
                with col_ve2:
                    val_pago_pe = st.number_input("Valor Pago (R$)", min_value=0.0, value=float(p_row_e['valor_pago'] or 0.0), step=500.0)
                with col_ve3:
                    val_adit_pe = st.number_input("Aditivos (R$)", min_value=0.0, value=float(p_row_e['aditivos'] or 0.0), step=100.0)
                    val_desc_pe = st.number_input("Descontos (R$)", min_value=0.0, value=float(p_row_e['descontos'] or 0.0), step=100.0)
                with col_ve4:
                    pag_opts = ["Em Aberto", "Parcialmente Pago", "Quitado", "Atrasado"]
                    pg_idx = pag_opts.index(p_row_e['status_pagamento']) if p_row_e['status_pagamento'] in pag_opts else 0
                    status_pag_pe = st.selectbox("Status do Pagamento", pag_opts, index=pg_idx)

                sub_edit_p = st.form_submit_button("🔄 Atualizar Projeto e Valores")
                if sub_edit_p:
                    success, msg = update_projeto(
                        p_edit_id, num_ent_e, emp_id_e, atividade_pe.strip(), uf_pe, mun_pe, num_proc_e,
                        'X' if lp_check_e else '', 'X' if li_check_e else '', 'X' if lo_check_e else '',
                        dt_emissao_pe.strftime('%Y-%m-%d'), dt_venc_pe.strftime('%Y-%m-%d'), None,
                        prazo_lic_pe, None, coment_pe,
                        val_proj_pe, val_pago_pe, val_adit_pe, val_desc_pe, status_pag_pe
                    )
                    if success:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

    # --- 4. REMOVER PROJETO ---
    with tab_del_p:
        st.markdown("### 🗑️ Remover Projeto")
        df_projs_del = list_projetos()
        if df_projs_del.empty:
            st.info("Nenhum projeto cadastrado.")
        else:
            p_del_id = st.selectbox("Selecione o Projeto a ser Removido:", df_projs_del['id'].tolist(), format_func=lambda x: f"ID #{x} - {df_projs_del[df_projs_del['id']==x]['cliente_nome'].values[0]} ({df_projs_del[df_projs_del['id']==x]['atividade'].values[0]})")
            st.warning("⚠️ **Atenção:** Esta operação excluirá o registro do projeto e todo o histórico financeiro associado.")
            
            if st.button("❌ Confirmar Exclusão do Projeto"):
                success, msg = delete_projeto(p_del_id)
                if success:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)


# ==============================================================================
# MÓDULO 4: RELATÓRIOS & EXPORTAÇÃO DE DADOS
# ==============================================================================
elif modulo == "📄 Relatórios & Exportação":
    st.markdown("""
        <div class="siga-header">
            <h1>📄 Relatórios e Gerador de Documentos</h1>
            <p>Emissão de relatórios de licenciamento ambiental e exportação de dados para Excel e CSV.</p>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📊 Gerar Relatório Personalizado")
    
    df_p_all = list_projetos()
    
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        filtro_mun = st.multiselect("Filtrar por Município(s):", options=sorted(df_p_all['municipio'].dropna().unique()))
    with col_r2:
        filtro_status_pag = st.multiselect("Filtrar por Status de Pagamento:", options=sorted(df_p_all['status_pagamento'].dropna().unique()))
        
    df_filtered_rep = df_p_all.copy()
    if filtro_mun:
        df_filtered_rep = df_filtered_rep[df_filtered_rep['municipio'].isin(filtro_mun)]
    if filtro_status_pag:
        df_filtered_rep = df_filtered_rep[df_filtered_rep['status_pagamento'].isin(filtro_status_pag)]
        
    st.write(f"**Registros encontrados ({len(df_filtered_rep)}):**")
    st.dataframe(df_filtered_rep, use_container_width=True, hide_index=True)
    
    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        csv_data = df_filtered_rep.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Baixar Relatório em CSV",
            data=csv_data,
            file_name=f"relatorio_siga_mt_{datetime.date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    with col_exp2:
        st.info("💡 É possível salvar como PDF imprimindo a página diretamente no seu navegador (`Ctrl + P`).")


# ==============================================================================
# MÓDULO 5: CONFIGURAÇÕES & BASE DE DADOS
# ==============================================================================
elif modulo == "⚙️ Configurações":
    st.markdown("""
        <div class="siga-header">
            <h1>⚙️ Configurações e Manutenção do Sistema</h1>
            <p>Gerenciamento da base de dados e recarga dos dados originais da planilha Excel.</p>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 🔄 Recarregar Dados da Planilha Original")
    st.write("Esta opção permite recarregar e restaurar os dados iniciais do arquivo `Protege-BETA EMPREENDIMENTO ATUAL (2).xls` no banco SQLite.")
    
    if st.button("⚠️ Recarregar Dados do Excel Original"):
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM projetos")
        cursor.execute("DELETE FROM empresas")
        conn.commit()
        cargar_dados_iniciais_excel(conn)
        conn.close()
        st.success("Base de dados restaurada a partir da planilha inicial!")
        st.rerun()
