import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from streamlit_gsheets import GSheetsConnection

# -----------------------------------------------------------------------------
# CONFIGURAÇÃO DA PÁGINA
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Gestão de Licenciamento Ambiental",
    page_icon="",
    layout="wide"
)

# -----------------------------------------------------------------------------
# ESTILIZAÇÃO CSS PERSONALIZADA
# -----------------------------------------------------------------------------
custom_css = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .stCard {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        border: 1px solid #e5e7eb;
        margin-bottom: 16px;
    }

    [data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0f766e;
    }

    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        border: none;
        transition: all 0.2s ease-in-out;
    }

    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    }

    [data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# INTEGRAÇÃO COM GOOGLE SHEETS
# -----------------------------------------------------------------------------
def obter_conexao():
    return st.connection("gsheets", type=GSheetsConnection)

def carregar_dados_gsheets():
    """Carrega dados das abas do Google Sheets para o Session State."""
    conn = obter_conexao()
    try:
        # Carregar Empresas
        df_emp = conn.read(worksheet="Empresas", ttl=0)
        if df_emp is not None and not df_emp.empty:
            st.session_state.empresas = df_emp.to_dict(orient="records")
        else:
            st.session_state.empresas = []

        # Carregar Tipos de Licenças
        df_lic = conn.read(worksheet="Tipos_Licencas", ttl=0)
        if df_lic is not None and not df_lic.empty:
            st.session_state.tipos_licencas = df_lic.to_dict(orient="records")
        else:
            st.session_state.tipos_licencas = []

        # Carregar Projetos
        df_proj = conn.read(worksheet="Projetos", ttl=0)
        if df_proj is not None and not df_proj.empty:
            df_proj["data_emissao"] = pd.to_datetime(df_proj["data_emissao"]).dt.date
            df_proj["data_vencimento"] = pd.to_datetime(df_proj["data_vencimento"]).dt.date
            st.session_state.projetos = df_proj.to_dict(orient="records")
        else:
            st.session_state.projetos = []
            
        return True
    except Exception as e:
        st.error(f"Erro ao ler dados do Google Sheets: {e}")
        return False

def salvar_dados_gsheets():
    """Salva todo o estado da aplicação de volta no Google Sheets."""
    conn = obter_conexao()
    try:
        # Salvar Empresas
        df_emp = pd.DataFrame(st.session_state.empresas)
        conn.update(worksheet="Empresas", data=df_emp)

        # Salvar Tipos Licenças
        df_lic = pd.DataFrame(st.session_state.tipos_licencas)
        conn.update(worksheet="Tipos_Licencas", data=df_lic)

        # Salvar Projetos
        df_proj = pd.DataFrame(st.session_state.projetos)
        if not df_proj.empty:
            df_proj["data_emissao"] = df_proj["data_emissao"].astype(str)
            df_proj["data_vencimento"] = df_proj["data_vencimento"].astype(str)
        conn.update(worksheet="Projetos", data=df_proj)

        return True
    except Exception as e:
        st.error(f"Erro ao salvar dados no Google Sheets: {e}")
        return False

# -----------------------------------------------------------------------------
# INICIALIZAÇÃO DE SESSÃO E CARREGAMENTO
# -----------------------------------------------------------------------------
if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = None

if "dados_carregados" not in st.session_state:
    st.session_state.dados_carregados = False

# -----------------------------------------------------------------------------
# AUTENTICAÇÃO E LOGIN
# -----------------------------------------------------------------------------
USUARIOS_VALIDOS = {
    "laura_nazario": "123456",
    "admin": "admin123"
}

if st.session_state.usuario_logado is None:
    st.markdown("<h1 style='text-align: center;'>Sistema de Licenciamento Ambiental</h1>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.subheader("🔑 Autenticação de Usuário")
        with st.form("form_login"):
            usuario_input = st.text_input("Usuário").strip().lower()
            senha_input = st.text_input("Senha", type="password")
            btn_login = st.form_submit_button("Acessar Sistema")

            if btn_login:
                if usuario_input in USUARIOS_VALIDOS and USUARIOS_VALIDOS[usuario_input] == senha_input:
                    st.session_state.usuario_logado = usuario_input
                    with st.spinner("Sincronizando com Google Sheets..."):
                        carregar_dados_gsheets()
                        st.session_state.dados_carregados = True
                    st.success(f"Bem-vindo(a), {usuario_input}!")
                    st.rerun()
                else:
                    st.error("Usuário ou senha inválidos.")
    st.stop()

# Autosave simplificado
def autosave_if_laura():
    if st.session_state.usuario_logado == "laura_nazario":
        salvar_dados_gsheets()

# -----------------------------------------------------------------------------
# BARRA LATERAL (NAV E STATUS)
# -----------------------------------------------------------------------------
with st.sidebar:
    st.image("logo_file.jpg", use_container_width=True)
    st.markdown(f"👤 **Usuário:** `{st.session_state.usuario_logado}`")
    
    if st.session_state.usuario_logado == "laura_nazario":
        st.success("🤖 **Sincronização Ativa**: Conectado ao Google Sheets.")
    
    if st.button("🚪 Sair / Logoff"):
        if st.session_state.usuario_logado == "laura_nazario":
            salvar_dados_gsheets()
            st.toast("💾 Dados salvos no Google Sheets!")
        st.session_state.usuario_logado = None
        st.session_state.dados_carregados = False
        st.rerun()

    st.markdown("---")
    
    menu_principal = st.radio(
        "Navegação Menu Principal:",
        [
            "🏢 Cadastros de Empresas",
            "📜 Regras de Licenças & Docs",
            "🚀 Projetos de Licenciamento",
            "📊 Relatórios Gerenciais",
            "⏰ Controle de Vencimentos e Alertas",
            "📁 Sincronização Google Sheets / Excel"
        ]
    )

# -----------------------------------------------------------------------------
# 1. CADASTRO DE EMPRESAS
# -----------------------------------------------------------------------------
if menu_principal == "🏢 Cadastros de Empresas":
    st.title("Cadastro de Empresas")
    sub_menu = st.radio("Selecione uma ação:", ["Inserir Empresa", "Editar Empresa", "Remover Empresa"], horizontal=True)

    if sub_menu == "Inserir Empresa":
        st.subheader("Cadastrar Nova Empresa")
        with st.form("form_add_empresa"):
            nome = st.text_input("Nome / Razão Social")
            cnpj = st.text_input("CNPJ")
            contato = st.text_input("E-mail de Contato")
            submitted = st.form_submit_button("Salvar Empresa")
            
            if submitted:
                if nome:
                    novo_id = max([e["id"] for e in st.session_state.empresas], default=0) + 1
                    st.session_state.empresas.append({"id": novo_id, "nome": nome, "cnpj": cnpj, "contato": contato})
                    autosave_if_laura()
                    st.success(f"Empresa '{nome}' cadastrada!")
                    st.rerun()
                else:
                    st.error("O campo Nome/Razão Social é obrigatório.")

    elif sub_menu == "Editar Empresa":
        st.subheader("Editar Dados da Empresa")
        if st.session_state.empresas:
            empresa_sel_nome = st.selectbox("Selecione a empresa:", [e["nome"] for e in st.session_state.empresas])
            empresa_dict = next(e for e in st.session_state.empresas if e["nome"] == empresa_sel_nome)
            
            with st.form("form_edit_empresa"):
                novo_nome = st.text_input("Nome / Razão Social", value=empresa_dict["nome"])
                novo_cnpj = st.text_input("CNPJ", value=empresa_dict["cnpj"])
                novo_contato = st.text_input("E-mail de Contato", value=empresa_dict["contato"])
                submitted = st.form_submit_button("Atualizar Empresa")
                
                if submitted:
                    empresa_dict["nome"] = novo_nome
                    empresa_dict["cnpj"] = novo_cnpj
                    empresa_dict["contato"] = novo_contato
                    autosave_if_laura()
                    st.success("Dados atualizados!")
                    st.rerun()

    elif sub_menu == "Remover Empresa":
        st.subheader("Remover Empresa")
        if st.session_state.empresas:
            empresa_sel_nome = st.selectbox("Selecione a empresa para remover:", [e["nome"] for e in st.session_state.empresas])
            empresa_dict = next(e for e in st.session_state.empresas if e["nome"] == empresa_sel_nome)
            
            if st.button("Confirmar Exclusão"):
                st.session_state.empresas = [e for e in st.session_state.empresas if e["id"] != empresa_dict["id"]]
                st.session_state.projetos = [p for p in st.session_state.projetos if p["empresa_id"] != empresa_dict["id"]]
                autosave_if_laura()
                st.success("Empresa removida com sucesso!")
                st.rerun()

    st.markdown("---")
    st.subheader("Empresas Cadastradas")
    if st.session_state.empresas:
        st.dataframe(pd.DataFrame(st.session_state.empresas), use_container_width=True)

# -----------------------------------------------------------------------------
# 2. REGRAS DE LICENÇAS E DOCS
# -----------------------------------------------------------------------------
elif menu_principal == "📜 Regras de Licenças & Docs":
    st.title("📜 Tipos de Licença, Prazos e Documentos")
    sub_menu_lic = st.radio("Selecione uma ação:", ["Inserir Licença", "Editar Licença"], horizontal=True)

    if sub_menu_lic == "Inserir Licença":
        st.subheader("Adicionar Novo Tipo de Licença")
        with st.form("form_add_licenca"):
            sigla = st.text_input("Sigla (ex: LP, LI, LO, LAS)").upper()
            nome_lic = st.text_input("Nome Completo")
            prazo_dias = st.number_input("Prazo Padrão (dias)", min_value=1, value=365)
            
            st.markdown("**Exigência de Documentos:**")
            doc_admin = st.checkbox("Documentos Administrativos")
            doc_tecnico = st.checkbox("Documentos Técnicos")
            
            submitted = st.form_submit_button("Salvar Tipo de Licença")
            if submitted and sigla and nome_lic:
                # Monta a lista de tipos de documentos selecionados
                docs_selecionados = []
                if doc_admin:
                    docs_selecionados.append("Administrativos")
                if doc_tecnico:
                    docs_selecionados.append("Técnicos")
                
                docs_str = ", ".join(docs_selecionados) if docs_selecionados else "Nenhum"

                st.session_state.tipos_licencas.append({
                    "sigla": sigla, 
                    "nome": nome_lic, 
                    "prazo_padrao_dias": prazo_dias, 
                    "documentos": docs_str
                })
                autosave_if_laura()
                st.success(f"Tipo '{sigla}' cadastrado!")
                st.rerun()

    elif sub_menu_lic == "Editar Licença":
        st.subheader("Editar Regras de Licenças e Docs")
        if st.session_state.tipos_licencas:
            licenca_sel_sigla = st.selectbox("Selecione a licença para editar:", [t["sigla"] for t in st.session_state.tipos_licencas])
            licenca_dict = next(t for t in st.session_state.tipos_licencas if t["sigla"] == licenca_sel_sigla)

            # Verifica o que já estava salvo para marcar os checkboxes
            docs_atuais = str(licenca_dict.get("documentos", ""))
            init_admin = "Administrativos" in docs_atuais
            init_tecnico = "Técnicos" in docs_atuais

            with st.form("form_edit_licenca"):
                nova_sigla = st.text_input("Sigla", value=licenca_dict["sigla"]).upper()
                novo_nome_lic = st.text_input("Nome Completo", value=licenca_dict["nome"])
                novo_prazo_dias = st.number_input("Prazo Padrão (dias)", min_value=1, value=int(licenca_dict["prazo_padrao_dias"]))
                
                st.markdown("**Exigência de Documentos:**")
                edit_admin = st.checkbox("Documentos Administrativos", value=init_admin)
                edit_tecnico = st.checkbox("Documentos Técnicos", value=init_tecnico)

                submitted_edit = st.form_submit_button("Atualizar Tipo de Licença")
                if submitted_edit:
                    docs_editados = []
                    if edit_admin:
                        docs_editados.append("Administrativos")
                    if edit_tecnico:
                        docs_editados.append("Técnicos")
                    
                    licenca_dict["sigla"] = nova_sigla
                    licenca_dict["nome"] = novo_nome_lic
                    licenca_dict["prazo_padrao_dias"] = novo_prazo_dias
                    licenca_dict["documentos"] = ", ".join(docs_editados) if docs_editados else "Nenhum"
                    
                    autosave_if_laura()
                    st.success("Licença atualizada com sucesso!")
                    st.rerun()

    st.markdown("---")
    st.subheader("Tipos Configurados")
    if st.session_state.tipos_licencas:
        st.dataframe(pd.DataFrame(st.session_state.tipos_licencas), use_container_width=True)

# -----------------------------------------------------------------------------
# 3. PROJETOS DE LICENCIAMENTO
# -----------------------------------------------------------------------------
elif menu_principal == "🚀 Projetos de Licenciamento":
    st.title("🚀 Gestão de Projetos de Licenciamento")

    if not st.session_state.empresas:
        st.warning("Cadastre ao menos uma Empresa antes de criar projetos.")
    else:
        sub_menu_proj = st.radio("Selecione uma ação:", ["Inserir Projeto", "Editar Projeto"], horizontal=True)
        empresa_map = {e["nome"]: e["id"] for e in st.session_state.empresas}
        empresa_id_to_nome = {e["id"]: e["nome"] for e in st.session_state.empresas}
        tipos_siglas = [t["sigla"] for t in st.session_state.tipos_licencas] if st.session_state.tipos_licencas else ["Outro"]

        if sub_menu_proj == "Inserir Projeto":
            empresa_nome_sel = st.selectbox("Selecione a Empresa (Cliente):", list(empresa_map.keys()))
            empresa_id_sel = empresa_map[empresa_nome_sel]

            with st.form("form_novo_projeto"):
                col1, col2 = st.columns(2)
                with col1:
                    nome_proj = st.text_input("Nome do Projeto")
                    tipo_lic = st.selectbox("Tipo de Licença", tipos_siglas)
                    valor_proj = st.number_input("Valor do Projeto (R$)", min_value=0.0, value=5000.0, step=500.0)
                
                with col2:
                    dt_emissao = st.date_input("Data de Emissão", format="DD/MM/YYYY")
                    prazo_default = next((t["prazo_padrao_dias"] for t in st.session_state.tipos_licencas if t["sigla"] == tipo_lic), 365)
                    dt_vencimento = st.date_input("Data de Vencimento", value=dt_emissao + timedelta(days=int(prazo_default)), format="DD/MM/YYYY")
                    status_proj = st.selectbox("Status Atual", ["Em andamento", "Aprovado", "Pendente Doc", "Cancelado"])

                sub_proj = st.form_submit_button("Cadastrar Projeto")
                if sub_proj:
                    novo_id_p = max([p["id"] for p in st.session_state.projetos], default=100) + 1
                    st.session_state.projetos.append({
                        "id": novo_id_p,
                        "empresa_id": empresa_id_sel,
                        "nome_projeto": nome_proj,
                        "tipo_licenca": tipo_lic,
                        "valor": valor_proj,
                        "data_emissao": dt_emissao,
                        "data_vencimento": dt_vencimento,
                        "status": status_proj
                    })
                    autosave_if_laura()
                    st.success("Projeto cadastrado com sucesso!")
                    st.rerun()

        elif sub_menu_proj == "Editar Projeto":
            st.subheader("Editar Projeto Existente")
            if st.session_state.projetos:
                proj_map = {f"ID: {p['id']} - {p['nome_projeto']}": p["id"] for p in st.session_state.projetos}
                proj_sel = st.selectbox("Selecione o projeto:", list(proj_map.keys()))
                proj_id = proj_map[proj_sel]
                proj_dict = next(p for p in st.session_state.projetos if p["id"] == proj_id)

                empresa_atual = empresa_id_to_nome.get(proj_dict["empresa_id"], list(empresa_map.keys())[0])

                with st.form("form_edit_projeto"):
                    col1, col2 = st.columns(2)
                    with col1:
                        empresa_edit_nome = st.selectbox("Empresa (Cliente)", list(empresa_map.keys()), index=list(empresa_map.keys()).index(empresa_atual))
                        nome_proj_edit = st.text_input("Nome do Projeto", value=proj_dict["nome_projeto"])
                        tipo_index = tipos_siglas.index(proj_dict["tipo_licenca"]) if proj_dict["tipo_licenca"] in tipos_siglas else 0
                        tipo_lic_edit = st.selectbox("Tipo de Licença", tipos_siglas, index=tipo_index)
                        valor_proj_edit = st.number_input("Valor do Projeto (R$)", min_value=0.0, value=float(proj_dict["valor"]), step=500.0)

                    with col2:
                        d_emissao = proj_dict["data_emissao"] if isinstance(proj_dict["data_emissao"], datetime) or hasattr(proj_dict["data_emissao"], "year") else pd.to_datetime(proj_dict["data_emissao"]).date()
                        d_venc = proj_dict["data_vencimento"] if isinstance(proj_dict["data_vencimento"], datetime) or hasattr(proj_dict["data_vencimento"], "year") else pd.to_datetime(proj_dict["data_vencimento"]).date()

                        dt_emissao_edit = st.date_input("Data de Emissão", value=d_emissao, format="DD/MM/YYYY")
                        dt_vencimento_edit = st.date_input("Data de Vencimento", value=d_venc, format="DD/MM/YYYY")
                        status_options = ["Em andamento", "Aprovado", "Pendente Doc", "Cancelado"]
                        status_index = status_options.index(proj_dict["status"]) if proj_dict["status"] in status_options else 0
                        status_proj_edit = st.selectbox("Status Atual", status_options, index=status_index)

                    submitted_edit_p = st.form_submit_button("Atualizar Projeto")
                    if submitted_edit_p:
                        proj_dict["empresa_id"] = empresa_map[empresa_edit_nome]
                        proj_dict["nome_projeto"] = nome_proj_edit
                        proj_dict["tipo_licenca"] = tipo_lic_edit
                        proj_dict["valor"] = valor_proj_edit
                        proj_dict["data_emissao"] = dt_emissao_edit
                        proj_dict["data_vencimento"] = dt_vencimento_edit
                        proj_dict["status"] = status_proj_edit
                        autosave_if_laura()
                        st.success("Projeto atualizado com sucesso!")
                        st.rerun()

    st.markdown("---")
    st.subheader("Projetos Cadastrados")
    if st.session_state.projetos:
        df_p_show = pd.DataFrame(st.session_state.projetos)
        df_p_show["data_emissao"] = pd.to_datetime(df_p_show["data_emissao"]).dt.strftime("%d/%m/%Y")
        df_p_show["data_vencimento"] = pd.to_datetime(df_p_show["data_vencimento"]).dt.strftime("%d/%m/%Y")
        st.dataframe(df_p_show, use_container_width=True)

# -----------------------------------------------------------------------------
# 4. RELATÓRIOS GERENCIAIS
# -----------------------------------------------------------------------------
elif menu_principal == "📊 Relatórios Gerenciais":
    st.title("📊 Relatórios e Indicadores")

    if st.session_state.projetos and st.session_state.empresas:
        df_proj = pd.DataFrame(st.session_state.projetos)
        df_emp = pd.DataFrame(st.session_state.empresas)
        df_completo = pd.merge(df_proj, df_emp, left_on="empresa_id", right_on="id", suffixes=("_proj", "_emp"))

        m1, m2, m3 = st.columns(3)
        m1.metric("Total de Projetos", len(df_completo))
        m2.metric("Valor Total Contratado", f"R$ {df_completo['valor'].sum():,.2f}")
        m3.metric("Empresas Atendidas", df_completo["nome"].nunique())

        df_completo["data_emissao"] = pd.to_datetime(df_completo["data_emissao"]).dt.strftime("%d/%m/%Y")
        df_completo["data_vencimento"] = pd.to_datetime(df_completo["data_vencimento"]).dt.strftime("%d/%m/%Y")

        st.dataframe(df_completo[["id_proj", "nome_projeto", "nome", "tipo_licenca", "data_emissao", "data_vencimento", "valor", "status"]], use_container_width=True)

# -----------------------------------------------------------------------------
# 5. CONTROLE DE VENCIMENTOS E ALERTAS
# -----------------------------------------------------------------------------
elif menu_principal == "⏰ Controle de Vencimentos e Alertas":
    st.title("⏰ Controle de Vencimento e Alertas")
    
    if st.session_state.projetos and st.session_state.empresas:
        df_p = pd.DataFrame(st.session_state.projetos)
        df_e = pd.DataFrame(st.session_state.empresas)
        df_alertas = pd.merge(df_p, df_e, left_on="empresa_id", right_on="id")

        dias_antecedencia = st.number_input("Filtrar licenças com vencimento em até (dias):", min_value=1, max_value=365, value=30, step=1)
        hoje = datetime.now().date()
        
        df_alertas["dias_para_vencer"] = df_alertas["data_vencimento"].apply(
            lambda x: (x - hoje).days if isinstance(x, datetime) or hasattr(x, 'days') else (pd.to_datetime(x).date() - hoje).days
        )
        df_filtrado = df_alertas[df_alertas["dias_para_vencer"] <= dias_antecedencia].copy()

        if not df_filtrado.empty:
            df_filtrado["data_vencimento_fmt"] = pd.to_datetime(df_filtrado["data_vencimento"]).dt.strftime("%d/%m/%Y")
            st.dataframe(df_filtrado[["nome_projeto", "nome", "tipo_licenca", "data_vencimento_fmt", "dias_para_vencer", "contato"]], use_container_width=True)

            st.markdown("---")
            st.subheader("📧 Enviar Alerta por E-mail")
            with st.form("form_envio_email"):
                email_destino = st.text_input("E-mail de destino do Alerta:")
                assunto_email = st.text_input("Assunto do E-mail", value="Alerta de Vencimento de Licenças Ambientais")
                enviar_btn = st.form_submit_button("Enviar Alerta")

                if enviar_btn:
                    if email_destino:
                        # Exemplo conceitual de estrutura de e-mail registrada localmente ou via SMTP
                        st.success(f"Alerta gerado com sucesso para {email_destino}! ({len(df_filtrado)} licença(s) no relatório)")
                    else:
                        st.error("Por favor, insira um e-mail válido.")
        else:
            st.info("Nenhuma licença com vencimento dentro do período especificado.")

# -----------------------------------------------------------------------------
# 6. SINCRONIZAÇÃO COM GOOGLE SHEETS
# -----------------------------------------------------------------------------
elif menu_principal == "📁 Sincronização Google Sheets / Excel":
    st.title("📁 Gerenciamento de Dados (Google Sheets)")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 Recarregar Dados do Google Sheets"):
            if carregar_dados_gsheets():
                st.success("Dados recarregados da nuvem!")
                st.rerun()

    with col2:
        if st.button("💾 Salvar Alterações na Nuvem Agora"):
            if salvar_dados_gsheets():
                st.success("Planilha no Google Sheets atualizada com sucesso!")

    st.markdown("---")
    st.subheader("📥 Exportação Manual Local")
    
    # Exportação para download local em Excel (backup)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(st.session_state.empresas).to_excel(writer, sheet_name="Empresas", index=False)
        pd.DataFrame(st.session_state.tipos_licencas).to_excel(writer, sheet_name="Tipos_Licencas", index=False)
        pd.DataFrame(st.session_state.projetos).to_excel(writer, sheet_name="Projetos", index=False)
    output.seek(0)

    st.download_button(
        label="💾 Baixar Cópia em Excel (.xlsx)",
        data=output,
        file_name=f"backup_licenciamento_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
