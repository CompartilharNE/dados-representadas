"""
app.py — Interface Streamlit
Compartilhar NE — Dados Representadas
Rodar: streamlit run app.py --server.port 8501
"""
import streamlit as st
import pandas as pd
import datetime
import base64
import banco
import gerar_excel
import importar as imp
try:
    import extra_streamlit_components as stx
    _COOKIES_OK = True
except ImportError:
    _COOKIES_OK = False

# ── CONFIGURAÇÃO DA PÁGINA ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dados Representadas — Compartilhar NE",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Garantir banco criado
banco.criar_banco()

# ── ESTILOS ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    [data-testid="stSidebar"] { background-color: #f5f5f0; }
    [data-testid="stSidebar"] * { color: rgba(39,42,74,0.75) !important; }
    [data-testid="stSidebar"] .sidebar-title { color: #c62621 !important; font-size: 1.1rem; font-weight: 600; }
    .metric-box { background: #f8f9fa; border-radius: 8px; padding: 12px 16px; text-align: center; }
    .metric-label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-val   { font-size: 26px; font-weight: 600; margin-top: 2px; }
    .green { color: #2e7d32; }
    .red   { color: #c62828; }
    .blue  { color: #1565c0; }
    .section-header { font-size: 13px; font-weight: 600; color: #444; margin: 8px 0 4px; }
    div[data-testid="stDownloadButton"] button { background-color: #1F4E79; color: white; font-weight: 600; width: 100%; }
    [data-testid="stSidebar"] div[data-testid="stButton"] button {
        background: transparent !important;
        border: none !important;
        color: #e63946 !important;
        font-size: 13px !important;
        font-weight: 700 !important;
        padding: 0 4px !important;
        box-shadow: none !important;
        text-align: left !important;
        width: auto !important;
    }
    [data-testid="stSidebar"] div[data-testid="stButton"] button:hover {
        color: #c1121f !important;
        background: transparent !important;
        text-decoration: underline !important;
    }
</style>
""", unsafe_allow_html=True)


# ── COOKIE MANAGER ───────────────────────────────────────────────────────────
@st.cache_resource
def _get_cookie_manager():
    return stx.CookieManager(key="dr_cm") if _COOKIES_OK else None

_cm = _get_cookie_manager() if _COOKIES_OK else None

def _cookie_get(name):
    try:
        return _cm.get(name) if _cm else None
    except Exception:
        return None

def _cookie_set(name, value, days=30):
    try:
        if _cm:
            exp = datetime.datetime.now() + datetime.timedelta(days=days)
            _cm.set(name, value, expires_at=exp)
    except Exception:
        pass

def _cookie_delete(name):
    try:
        if _cm:
            _cm.delete(name)
    except Exception:
        pass


# ── AUTENTICAÇÃO ──────────────────────────────────────────────────────────────
def _check_admin(username, senha):
    """Verifica credenciais do admin via st.secrets."""
    try:
        adm_user = st.secrets["auth"]["admin_username"]
        adm_pass = st.secrets["auth"]["admin_password"]
        return username.strip().lower() == adm_user.strip().lower() and senha == adm_pass
    except Exception:
        return False

def _fazer_login(username, senha):
    """Tenta login como admin (secrets) ou usuário (banco). Retorna perfil ou None."""
    if _check_admin(username, senha):
        return "admin"
    usuario = banco.verificar_login(username, senha)
    if usuario:
        return usuario["perfil"]
    return None

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.perfil    = None
    st.session_state.username  = None

# Auto-login via cookie "lembrar-me"
if not st.session_state.logged_in and _cm:
    _c_user = _cookie_get("dr_u")
    _c_pass = _cookie_get("dr_p")
    if _c_user and _c_pass:
        try:
            _c_pass_dec = base64.b64decode(_c_pass.encode()).decode()
            _c_perfil   = _fazer_login(_c_user, _c_pass_dec)
            if _c_perfil:
                st.session_state.logged_in = True
                st.session_state.perfil    = _c_perfil
                st.session_state.username  = _c_user
                st.rerun()
        except Exception:
            _cookie_delete("dr_u")
            _cookie_delete("dr_p")

if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        try:
            st.image("compartilhar_logo.svg", use_container_width=True)
        except Exception:
            pass
        st.markdown("### Dados Representadas")
        st.caption("Acesso restrito — Compartilhar NE")
        with st.form("login_form"):
            username_inp = st.text_input("Usuário")
            senha_inp    = st.text_input("Senha", type="password")
            lembrar      = st.checkbox("Lembrar login e senha neste dispositivo", value=False)
            entrar       = st.form_submit_button("Entrar", use_container_width=True, type="primary")
            if entrar:
                perfil = _fazer_login(username_inp, senha_inp)
                if perfil:
                    st.session_state.logged_in = True
                    st.session_state.perfil    = perfil
                    st.session_state.username  = username_inp.strip().lower()
                    if lembrar:
                        _cookie_set("dr_u", username_inp.strip().lower(), days=30)
                        _cookie_set("dr_p", base64.b64encode(senha_inp.encode()).decode(), days=30)
                    else:
                        _cookie_delete("dr_u")
                        _cookie_delete("dr_p")
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")
    st.stop()

# ── NAVEGAÇÃO SIDEBAR ─────────────────────────────────────────────────────────
_is_admin = st.session_state.perfil == "admin"

with st.sidebar:
    st.image("compartilhar_logo.svg", use_container_width=True)
    st.markdown('<div class="sidebar-title" style="margin-top:6px">📊 Dados Representadas</div>', unsafe_allow_html=True)
    st.markdown('<div style="font-size:11px;color:rgba(255,255,255,0.4);margin-bottom:12px">Compartilhar NE</div>', unsafe_allow_html=True)

    if _is_admin:
        opcoes_nav = [
            "🗂️ Gerar Relatório", "🏭 Fábricas", "🏪 Redes / Clientes",
            "🏬 Lojas", "📦 Produtos", "🔢 Códigos da Rede", "💰 Tabela de Preço",
            "📥 Importar Faturamento", "👥 Usuários",
        ]
    else:
        opcoes_nav = ["🗂️ Gerar Relatório"]

    pagina = st.radio("Navegação", opcoes_nav, label_visibility="collapsed")

    st.markdown("---")
    st.markdown(
        f'<div style="font-size:11px;color:rgba(39,42,74,0.5)">'
        f'👤 {st.session_state.username} '
        f'<span style="background:#e63946;color:#fff;padding:1px 6px;border-radius:4px;font-size:10px">'
        f'{st.session_state.perfil}</span></div>',
        unsafe_allow_html=True,
    )
    if st.button("↩ Sair", use_container_width=False):
        _cookie_delete("dr_u")
        _cookie_delete("dr_p")
        st.session_state.logged_in = False
        st.session_state.perfil    = None
        st.session_state.username  = None
        st.rerun()
    st.markdown('<div style="font-size:10px;color:rgba(255,255,255,0.3)">Banco: Supabase PostgreSQL</div>',
                unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: GERAR RELATÓRIO
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "🗂️ Gerar Relatório":
    st.title("Gerar Relatório")
    st.caption("Selecione a fábrica, as redes e o período para gerar o Excel de Dados Representadas.")

    fabricas = banco.listar_fabricas()
    redes    = banco.listar_redes()
    nomes_fab  = [f["nome"] for f in fabricas]
    nomes_rede = [r["nome"] for r in redes]

    col1, col2 = st.columns(2)
    with col1:
        fab_sel = st.selectbox("Fábrica", [""] + nomes_fab, key="sel_fab")
    with col2:
        redes_sel = st.multiselect("Rede(s) / Cliente(s)", nomes_rede,
                                   default=[],
                                   key="sel_rede")

    col3, col4 = st.columns(2)
    with col3:
        d_ini = st.date_input("Data início", value=datetime.date(datetime.date.today().year, 1, 1), format="DD/MM/YYYY")
    with col4:
        d_fim = st.date_input("Data fim", value=datetime.date.today(), format="DD/MM/YYYY")

    st.markdown("")
    gerar = st.button("📊 Gerar Relatório", type="primary", use_container_width=True)

    if gerar:
        if not fab_sel or fab_sel == "":
            st.warning("Selecione uma fábrica.")
        elif not redes_sel:
            st.warning("Selecione pelo menos uma rede.")
        else:
            with st.spinner("Gerando relatório..."):
                excel_bytes, erro = gerar_excel.gerar_relatorio(
                    fab_nome=fab_sel,
                    redes_selecionadas=redes_sel,
                    data_inicio=str(d_ini),
                    data_fim=str(d_fim),
                )
            if erro:
                st.error(f"Erro: {erro}")
            else:
                st.success("Relatório gerado!")

                # Preview de métricas
                fab = banco.obter_fabrica_por_nome(fab_sel)
                m_col = st.columns(4)
                total_comp = total_ncomp = total_val = total_cat = 0
                for r_nome in redes_sel:
                    rede = banco.obter_rede_por_nome(r_nome)
                    if not rede:
                        continue
                    prods = banco.produtos_com_codigos(fab["id"], rede["id"])
                    vendas = banco.vendas_por_produto(fab["id"], rede["id"], str(d_ini), str(d_fim))
                    comp   = sum(1 for p in prods if vendas.get(p["id"], {}).get("valor_total", 0) > 0)
                    ncomp  = len(prods) - comp
                    val    = sum(v.get("valor_total", 0) for v in vendas.values())
                    total_cat   += len(prods)
                    total_comp  += comp
                    total_ncomp += ncomp
                    total_val   += val

                porc = round(total_comp / total_cat * 100, 1) if total_cat else 0
                with m_col[0]:
                    st.markdown(f'<div class="metric-box"><div class="metric-label">Cadastrados</div>'
                                f'<div class="metric-val blue">{total_cat}</div></div>', unsafe_allow_html=True)
                with m_col[1]:
                    st.markdown(f'<div class="metric-box"><div class="metric-label">Comprados</div>'
                                f'<div class="metric-val green">{total_comp}</div></div>', unsafe_allow_html=True)
                with m_col[2]:
                    st.markdown(f'<div class="metric-box"><div class="metric-label">Sem venda</div>'
                                f'<div class="metric-val red">{total_ncomp}</div></div>', unsafe_allow_html=True)
                with m_col[3]:
                    st.markdown(f'<div class="metric-box"><div class="metric-label">Positivação</div>'
                                f'<div class="metric-val">{porc}%</div></div>', unsafe_allow_html=True)

                st.markdown("")
                st.markdown(f"**Valor faturado total: R$ {total_val:,.2f}**")
                st.markdown("")

                nome_arquivo = (
                    f"Dados Representadas {fab_sel} - "
                    f"{', '.join(redes_sel)}.xlsx"
                )
                st.download_button(
                    label="⬇️ Baixar Excel",
                    data=excel_bytes,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                # Tabela preview
                if total_cat > 0:
                    st.markdown("---")
                    st.markdown("**Prévia dos itens**")
                    rows_prev = []
                    for r_nome in redes_sel:
                        rede = banco.obter_rede_por_nome(r_nome)
                        if not rede:
                            continue
                        prods  = banco.produtos_com_codigos(fab["id"], rede["id"])
                        vendas = banco.vendas_por_produto(fab["id"], rede["id"], str(d_ini), str(d_fim))
                        for p in prods:
                            vd = vendas.get(p["id"], {})
                            rows_prev.append({
                                "Rede":         r_nome,
                                "Cód. Fab":     p["codigo_fab"],
                                "Cód. Rede":    p["codigo_rede"],
                                "Produto":      p["nome"],
                                "Status":       "✅ Comprado" if vd.get("valor_total", 0) > 0 else "❌ Sem venda",
                                "Valor (R$)":   f"R$ {vd.get('valor_total',0):,.2f}" if vd.get("valor_total",0) > 0 else "—",
                                "Nº Lojas":     vd.get("n_lojas", "—"),
                            })
                    if rows_prev:
                        st.dataframe(pd.DataFrame(rows_prev), use_container_width=True, height=300)


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: FÁBRICAS
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "🏭 Fábricas":
    _c1, _c2 = st.columns([5, 1])
    with _c1:
        st.title("Fábricas")
    with _c2:
        if st.button("🔄 Atualizar", key="ref_fab"):
            st.rerun()
    st.caption("Fornecedores representados pela Compartilhar NE.")

    fabricas = banco.listar_fabricas()

    # Tabela
    df = pd.DataFrame([{
        "Nome":              f["nome"],
        "Nome no faturamento": f["nome_faturamento"],
        "Contato":           f["contato"] or "—",
        "Produtos":          banco.contar_produtos(f["id"]),
    } for f in fabricas])
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    with st.expander("➕ Adicionar / Editar fábrica"):
        nomes = ["Nova fábrica"] + [f["nome"] for f in fabricas]
        sel = st.selectbox("Selecionar para editar", nomes)
        fab_edit = next((f for f in fabricas if f["nome"] == sel), None)

        col1, col2 = st.columns(2)
        with col1:
            nome_in = st.text_input("Nome da fábrica", value=fab_edit["nome"] if fab_edit else "")
        with col2:
            nome_fat_in = st.text_input("Nome no faturamento",
                                        value=fab_edit["nome_faturamento"] if fab_edit else "")
        contato_in = st.text_input("Contato (opcional)", value=fab_edit["contato"] if fab_edit else "")

        # ── Logo da fábrica ───────────────────────────────────────────────────
        st.markdown("**Logo da fábrica** *(aparecerá na planilha ao lado da logo da rede)*")
        fab_logo_col1, fab_logo_col2 = st.columns([2, 3])
        with fab_logo_col1:
            fab_logo_atual = fab_edit.get("logo_b64", "") if fab_edit else ""
            if fab_logo_atual:
                st.image(base64.b64decode(fab_logo_atual), width=150, caption="Logo atual")
            else:
                st.caption("Nenhuma logo cadastrada")
        with fab_logo_col2:
            fab_logo_file = st.file_uploader(
                "Upload logo (PNG, JPG)", type=["png", "jpg", "jpeg"],
                key="fab_logo_up", help="Recomendado: fundo branco, aprox. 300×100 px"
            )
            if fab_logo_file:
                st.image(fab_logo_file, width=150, caption="Nova logo (prévia)")

        c1, c2 = st.columns([3, 1])
        with c1:
            if st.button("💾 Salvar", use_container_width=True):
                if nome_in and nome_fat_in:
                    banco.salvar_fabrica(nome_in, nome_fat_in, contato_in,
                                         fab_id=fab_edit["id"] if fab_edit else None)
                    # Salvar logo se foi enviada
                    if fab_logo_file and fab_edit:
                        fab_logo_b64 = base64.b64encode(fab_logo_file.getvalue()).decode()
                        banco.salvar_logo_fabrica(fab_edit["id"], fab_logo_b64)
                    elif fab_logo_file:
                        fab_nova = banco.obter_fabrica_por_nome(nome_in)
                        if fab_nova:
                            fab_logo_b64 = base64.b64encode(fab_logo_file.getvalue()).decode()
                            banco.salvar_logo_fabrica(fab_nova["id"], fab_logo_b64)
                    st.success("Salvo!")
                    st.rerun()
                else:
                    st.warning("Preencha nome e nome no faturamento.")
        with c2:
            if fab_edit and st.button("🗑️ Excluir", use_container_width=True):
                banco.deletar_fabrica(fab_edit["id"])
                st.success("Excluído.")
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: REDES / CLIENTES
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "🏪 Redes / Clientes":
    _c1, _c2 = st.columns([5, 1])
    with _c1:
        st.title("Redes / Clientes")
    with _c2:
        if st.button("🔄 Atualizar", key="ref_redes"):
            st.rerun()
    st.caption("Clientes onde os produtos são vendidos. O 'Filtro nome' é o texto que aparece no faturamento.")

    redes = banco.listar_redes()
    df = pd.DataFrame([{
        "Nome":           r["nome"],
        "Filtro no faturamento": r["filtro_nome"],
        "Estados":        r["estados"] or "Todos",
        "Excluir (palavras)": r["excluir_palavras"] or "—",
    } for r in redes])
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    with st.expander("➕ Adicionar / Editar rede"):
        nomes = ["Nova rede"] + [r["nome"] for r in redes]
        sel   = st.selectbox("Selecionar para editar", nomes)
        r_edit = next((r for r in redes if r["nome"] == sel), None)

        col1, col2 = st.columns(2)
        with col1:
            nome_in   = st.text_input("Nome da rede", value=r_edit["nome"] if r_edit else "")
            filtro_in = st.text_input("Filtro no faturamento (parte do nome do cliente)",
                                      value=r_edit["filtro_nome"] if r_edit else "")
        with col2:
            estados_in = st.text_input("Estados (separados por vírgula, ex: PE,PB,RN,AL)",
                                       value=r_edit["estados"] if r_edit else "")
            excluir_in = st.text_input("Palavras a excluir (separadas por vírgula)",
                                       value=r_edit["excluir_palavras"] if r_edit else "")

        # ── Logo da rede ──────────────────────────────────────────────────────
        st.markdown("**Logo da rede** *(aparecerá na planilha ao lado da logo da Fábrica)*")
        logo_col1, logo_col2 = st.columns([2, 3])
        with logo_col1:
            logo_atual = r_edit.get("logo_b64", "") if r_edit else ""
            if logo_atual:
                st.image(base64.b64decode(logo_atual), width=150, caption="Logo atual")
            else:
                st.caption("Nenhuma logo cadastrada")
        with logo_col2:
            logo_file = st.file_uploader(
                "Upload logo (PNG, JPG)", type=["png", "jpg", "jpeg"],
                key="rede_logo_up", help="Recomendado: fundo branco, aprox. 300×100 px"
            )
            if logo_file:
                st.image(logo_file, width=150, caption="Nova logo (prévia)")

        c1, c2 = st.columns([3, 1])
        with c1:
            if st.button("💾 Salvar rede", use_container_width=True):
                if nome_in and filtro_in:
                    banco.salvar_rede(nome_in, filtro_in, estados_in, excluir_in,
                                      rede_id=r_edit["id"] if r_edit else None)
                    # Salvar logo se foi enviada
                    if logo_file and r_edit:
                        logo_b64 = base64.b64encode(logo_file.getvalue()).decode()
                        banco.salvar_logo_rede(r_edit["id"], logo_b64)
                    elif logo_file:
                        # Busca o id recém-criado
                        rede_nova = banco.obter_rede_por_nome(nome_in)
                        if rede_nova:
                            logo_b64 = base64.b64encode(logo_file.getvalue()).decode()
                            banco.salvar_logo_rede(rede_nova["id"], logo_b64)
                    st.success("Salvo!")
                    st.rerun()
                else:
                    st.warning("Preencha nome e filtro.")
        with c2:
            if r_edit and st.button("🗑️ Excluir rede", use_container_width=True):
                try:
                    banco.deletar_rede(r_edit["id"])
                    st.success("Excluído.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao excluir: {e}")

    # ── Códigos de fornecedor por fábrica ─────────────────────────────────────
    st.markdown("---")
    st.subheader("🔑 Código de Fornecedor por Fábrica")
    st.caption(
        "Cada fábrica pode ter um código de fornecedor diferente para cada rede. "
        "Este código aparecerá na planilha gerada."
    )

    fabricas_all = banco.listar_fabricas()
    redes_cf     = banco.listar_redes()

    cf_col1, cf_col2 = st.columns(2)
    with cf_col1:
        rede_cf = st.selectbox("Rede / Cliente", [r["nome"] for r in redes_cf], key="cf_rede")
    rede_cf_obj = next((r for r in redes_cf if r["nome"] == rede_cf), None)

    if rede_cf_obj:
        cods_existentes = {
            c["fabrica_id"]: c["codigo"]
            for c in banco.listar_cods_forn_por_rede(rede_cf_obj["id"])
        }
        st.markdown(f"**Códigos cadastrados para {rede_cf}:**")
        rows_cf = []
        for fab in fabricas_all:
            cod = cods_existentes.get(fab["id"], "—")
            rows_cf.append({"Fábrica": fab["nome"], "Código Fornecedor": cod})
        st.dataframe(pd.DataFrame(rows_cf), use_container_width=True, hide_index=True)

        st.markdown("**Adicionar / Atualizar código:**")
        cf2, cf3, cf4 = st.columns([2, 2, 1])
        with cf2:
            fab_cf = st.selectbox("Fábrica", [f["nome"] for f in fabricas_all], key="cf_fab")
        with cf3:
            fab_cf_obj = next((f for f in fabricas_all if f["nome"] == fab_cf), None)
            val_atual  = cods_existentes.get(fab_cf_obj["id"], "") if fab_cf_obj else ""
            cod_in = st.text_input("Código do Fornecedor", value=val_atual, key="cf_cod")
        with cf4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("💾 Salvar", key="cf_salvar", use_container_width=True):
                if fab_cf_obj and cod_in.strip():
                    banco.salvar_cod_fornecedor(fab_cf_obj["id"], rede_cf_obj["id"], cod_in.strip())
                    st.success("Código salvo!")
                    st.rerun()
                else:
                    st.warning("Preencha o código.")



# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: LOJAS
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "🏬 Lojas":
    import io as _io

    _c1, _c2 = st.columns([5, 1])
    with _c1:
        st.title("Lojas")
    with _c2:
        if st.button("🔄 Atualizar", key="ref_lojas"):
            st.rerun()
    st.caption("Cadastre as lojas de cada rede. Após importar, o faturamento reconhece as lojas automaticamente sem precisar reimportar o arquivo.")

    redes_all = banco.listar_redes()

    # ── Importar planilha ────────────────────────────────────────────────────
    with st.expander("📥 Importar lojas de uma planilha (todas as redes)", expanded=True):
        st.markdown(
            "A planilha deve ter **3 colunas**: nome da loja (exatamente como aparece no faturamento), "
            "UF e nome da rede (igual ao cadastrado no sistema)."
        )

        template_df = pd.DataFrame({
            "Nome da Loja": ["Sendas Federação (Av. Vasco da Gama) (256)", "Sendas Madalena LJ 284"],
            "UF":           ["BA", "PE"],
            "Rede":         ["Assai BA", "Assai PE"],
        })
        buf_tpl = _io.BytesIO()
        template_df.to_excel(buf_tpl, index=False)
        st.download_button(
            "⬇️ Baixar modelo de planilha",
            data=buf_tpl.getvalue(),
            file_name="modelo_lojas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        arq_lojas = st.file_uploader("Arquivo (.xlsx ou .xls)", type=["xlsx", "xls"], key="up_lojas")

        if arq_lojas:
            try:
                df_raw = pd.read_excel(_io.BytesIO(arq_lojas.getvalue()))
                st.dataframe(df_raw.head(5), use_container_width=True, hide_index=True)

                colunas = list(df_raw.columns)
                c1, c2, c3 = st.columns(3)
                with c1:
                    col_nome = st.selectbox("Coluna: Nome da Loja", colunas,
                                            index=next((i for i, c in enumerate(colunas) if "nome" in c.lower() or "loja" in c.lower()), 0))
                with c2:
                    col_uf   = st.selectbox("Coluna: UF", colunas,
                                            index=next((i for i, c in enumerate(colunas) if "uf" in c.lower() or "estado" in c.lower()), min(1, len(colunas)-1)))
                with c3:
                    col_rede = st.selectbox("Coluna: Rede", colunas,
                                            index=next((i for i, c in enumerate(colunas) if "rede" in c.lower()), min(2, len(colunas)-1)))

                if st.button("📥 Importar lojas", type="primary", use_container_width=True):
                    n_ok, nao_enc, erros = banco.importar_lojas_planilha(df_raw, col_nome, col_uf, col_rede)
                    st.success(f"{n_ok} lojas importadas / atualizadas.")
                    if nao_enc:
                        st.warning(f"Redes não encontradas no sistema: {', '.join(nao_enc)}")
                    if erros:
                        st.error(f"{len(erros)} erros: {erros[:3]}")
                    st.rerun()
            except Exception as e:
                st.error(f"Erro ao ler planilha: {e}")

    st.markdown("---")

    # ── Listagem ─────────────────────────────────────────────────────────────
    col_filt, col_info = st.columns([3, 1])
    with col_filt:
        rede_filtro = st.selectbox(
            "Filtrar por rede",
            ["Todas as redes"] + [r["nome"] for r in redes_all],
            key="lojas_filt_rede",
        )
    rede_filt_obj = next((r for r in redes_all if r["nome"] == rede_filtro), None)
    lojas_all = banco.listar_lojas_com_rede(rede_id=rede_filt_obj["id"] if rede_filt_obj else None)

    with col_info:
        st.metric("Total de lojas", len(lojas_all))

    sem_uf   = [l for l in lojas_all if not l.get("estado")]
    sem_rede = [l for l in lojas_all if not l.get("rede_nome")]

    if sem_uf:
        st.warning(f"⚠️ {len(sem_uf)} loja(s) com UF não identificada — corrija manualmente abaixo.")
    if sem_rede:
        st.warning(f"⚠️ {len(sem_rede)} loja(s) sem rede associada.")

    ufs_br = ["","AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT",
              "PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO"]

    if lojas_all:
        header = st.columns([4, 1, 2, 1, 1])
        header[0].markdown("**Nome da loja**")
        header[1].markdown("**UF**")
        header[2].markdown("**Rede**")
        header[3].markdown("")
        header[4].markdown("")
        st.markdown('<hr style="margin:4px 0">', unsafe_allow_html=True)

        for loja in lojas_all:
            lid = loja["id"]
            c1, c2, c3, c4, c5 = st.columns([4, 1, 2, 1, 1])
            uf_label = loja.get("estado") or "?"
            uf_color = "#1565c0" if loja.get("estado") else "#e65100"
            c1.markdown(f'<span style="font-size:13px">{loja["nome_faturamento"]}</span>', unsafe_allow_html=True)
            c2.markdown(
                f'<span style="background:{"#e3f2fd" if loja.get("estado") else "#fff3e0"};' 
                f'color:{uf_color};padding:2px 8px;border-radius:6px;font-size:12px">{uf_label}</span>',
                unsafe_allow_html=True,
            )
            c3.markdown(f'<span style="font-size:13px;color:#555">{loja.get("rede_nome") or "—"}</span>', unsafe_allow_html=True)
            if c4.button("✏️", key=f"edit_loja_{lid}", help="Editar loja"):
                st.session_state[f"_edit_loja_{lid}"] = not st.session_state.get(f"_edit_loja_{lid}", False)
            if c5.button("🗑️", key=f"del_loja_{lid}", help="Remover loja"):
                banco.deletar_loja(lid)
                st.rerun()

            # Formulário inline de edição
            if st.session_state.get(f"_edit_loja_{lid}"):
                with st.form(f"form_edit_loja_{lid}"):
                    ec1, ec2, ec3 = st.columns([4, 1, 2])
                    novo_nome = ec1.text_input("Nome da loja", value=loja["nome_faturamento"], key=f"en_{lid}")
                    uf_idx = ufs_br.index(loja.get("estado") or "") if (loja.get("estado") or "") in ufs_br else 0
                    novo_uf  = ec2.selectbox("UF", ufs_br, index=uf_idx, key=f"eu_{lid}")
                    rede_nomes = [r["nome"] for r in redes_all]
                    rede_idx = rede_nomes.index(loja["rede_nome"]) if loja.get("rede_nome") in rede_nomes else 0
                    novo_rede = ec3.selectbox("Rede", rede_nomes, index=rede_idx, key=f"er_{lid}")
                    salvar_e, cancelar_e = st.columns(2)
                    if salvar_e.form_submit_button("💾 Salvar", use_container_width=True):
                        rede_obj_e = next((r for r in redes_all if r["nome"] == novo_rede), None)
                        if rede_obj_e and novo_nome.strip():
                            banco.atualizar_loja(lid, rede_obj_e["id"], novo_nome.strip(), novo_uf)
                            del st.session_state[f"_edit_loja_{lid}"]
                            st.rerun()
                    if cancelar_e.form_submit_button("Cancelar", use_container_width=True):
                        del st.session_state[f"_edit_loja_{lid}"]
                        st.rerun()
    else:
        st.info("Nenhuma loja cadastrada ainda.")

    st.markdown("---")

    # ── Adicionar manualmente ─────────────────────────────────────────────────
    with st.expander("➕ Adicionar loja manualmente"):
        ufs_br = ["AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT",
                  "PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO"]
        ma1, ma2, ma3, ma4 = st.columns([4, 1, 2, 1])
        with ma1:
            nome_man = st.text_input("Nome exato da loja (como aparece no faturamento)", key="man_nome")
        with ma2:
            uf_man   = st.selectbox("UF", ufs_br, key="man_uf")
        with ma3:
            rede_man = st.selectbox("Rede", [r["nome"] for r in redes_all], key="man_rede")
        with ma4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Adicionar", key="man_add", use_container_width=True):
                if nome_man.strip():
                    rede_man_obj = next((r for r in redes_all if r["nome"] == rede_man), None)
                    if rede_man_obj:
                        banco.salvar_loja_manual(rede_man_obj["id"], nome_man.strip(), uf_man)
                        st.success("Loja adicionada!")
                        st.rerun()
                else:
                    st.warning("Informe o nome da loja.")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: PRODUTOS
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "📦 Produtos":
    _c1, _c2 = st.columns([5, 1])
    with _c1:
        st.title("Produtos")
    with _c2:
        if st.button("🔄 Atualizar", key="ref_prod"):
            st.rerun()
    st.caption("Catálogo de produtos por fábrica.")

    fabricas = banco.listar_fabricas()
    nomes_fab = [f["nome"] for f in fabricas]
    fab_sel   = st.selectbox("Fábrica", nomes_fab)
    fab       = next((f for f in fabricas if f["nome"] == fab_sel), None)

    if fab:
        prods = banco.listar_produtos(fab["id"])
        if prods:
            df = pd.DataFrame([{
                "Cód. Fab":  p["codigo_fab"],
                "Produto":   p["nome"],
                "Família":   p["familia"] or "—",
                "Peso CX":   p["peso"] or "—",
                "Qtde CX":   p["qtde_cx"] or "—",
                "UND":       p["und"] or "—",
            } for p in prods])
            st.dataframe(df, use_container_width=True, hide_index=True, height=300)
            st.caption(f"{len(prods)} produtos cadastrados para {fab_sel}")
        else:
            st.info(f"Nenhum produto cadastrado para {fab_sel}. Importe o catálogo abaixo.")

        st.markdown("---")
        with st.expander("✏️ Cadastrar / Editar / Excluir produto", expanded=False):
            # Seletor para editar produto existente ou criar novo
            opcoes_prod = {"➕ Novo produto": None}
            if prods:
                for p in prods:
                    opcoes_prod[f"{p['codigo_fab']} — {p['nome']}"] = p
            prod_sel_key = st.selectbox("Selecione o produto", list(opcoes_prod.keys()), key="prod_edit_sel")
            p_edit = opcoes_prod[prod_sel_key]

            with st.form("form_produto"):
                fc1, fc2 = st.columns(2)
                cod_in  = fc1.text_input("Código *", value=p_edit["codigo_fab"] if p_edit else "")
                nome_in = fc2.text_input("Nome *",   value=p_edit["nome"] if p_edit else "")
                fc3, fc4, fc5, fc6 = st.columns(4)
                fam_in   = fc3.text_input("Família",  value=p_edit["familia"]  or "" if p_edit else "")
                peso_in  = fc4.text_input("Peso CX",  value=p_edit["peso"]     or "" if p_edit else "")
                qtd_in   = fc5.text_input("Qtde CX",  value=p_edit["qtde_cx"] or "" if p_edit else "")
                und_in   = fc6.text_input("UND",       value=p_edit["und"]      or "" if p_edit else "")

                sb1, sb2 = st.columns([3, 1])
                salvar  = sb1.form_submit_button("💾 Salvar produto", use_container_width=True)
                excluir = sb2.form_submit_button("🗑️ Excluir", use_container_width=True) if p_edit else False

                if salvar:
                    if cod_in and nome_in:
                        banco.salvar_produto(fab["id"], cod_in, nome_in, fam_in, peso_in, qtd_in, und_in, "")
                        st.success("Produto salvo!")
                        st.rerun()
                    else:
                        st.warning("Código e Nome são obrigatórios.")
                if excluir and p_edit:
                    try:
                        banco._run("DELETE FROM produtos WHERE id=%s", (p_edit["id"],))
                        st.success("Produto excluído.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro: {e} — verifique se há códigos de rede ou faturamento vinculados.")

    st.markdown("---")
    st.markdown("**Importar catálogo (.xlsx)**")
    st.caption("Envie qualquer planilha — o sistema detecta as colunas automaticamente e você confirma o mapeamento.")

    arq = st.file_uploader("Selecione o arquivo de catálogo", type=["xlsx", "xls"],
                            key="upload_catalogo")
    if arq is not None:
        _bytes_cat = arq.read()
        if _bytes_cat:
            st.session_state["_cache_catalogo"] = (arq.name, _bytes_cat)

    import io as _io
    _cat_cache = st.session_state.get("_cache_catalogo")
    if _cat_cache and fab:
        _arq_name, _arq_bytes = _cat_cache
        if arq is None:
            st.info(f"📄 Arquivo em cache: **{_arq_name}**")
        try:
            df_raw = pd.read_excel(_io.BytesIO(_arq_bytes), header=None)

            # Auto-detectar linha de cabeçalho (procura linha com palavras-chave)
            header_row = 0
            for i, row in df_raw.iterrows():
                vals = [str(v).lower() for v in row.values]
                joined = " ".join(vals)
                if any(kw in joined for kw in ["código", "cod", "produto", "nome", "descrição"]):
                    header_row = i
                    break
            df_cat = df_raw.copy()
            # Deduplica nomes de colunas (colunas vazias/repetidas viram _1, _2...)
            raw_cols = [str(c).strip() for c in df_raw.iloc[header_row]]
            seen = {}
            dedup_cols = []
            for c in raw_cols:
                if c in seen:
                    seen[c] += 1
                    dedup_cols.append(f"{c}_{seen[c]}" if c != "nan" else f"_col{seen[c]}")
                else:
                    seen[c] = 0
                    dedup_cols.append(c if c != "nan" else f"_col{len(dedup_cols)}")
            df_cat.columns = dedup_cols
            df_cat = df_cat.iloc[header_row + 1:].reset_index(drop=True)
            df_cat = df_cat.dropna(how="all")
            # Remove colunas sem nome útil
            df_cat = df_cat[[c for c in df_cat.columns if not c.startswith("_col")]]

            colunas = ["(ignorar)"] + list(df_cat.columns)

            # Função para sugerir coluna por palavras-chave
            def _sugerir(keywords):
                for col in df_cat.columns:
                    c = str(col).lower()
                    if any(kw in c for kw in keywords):
                        return col
                return "(ignorar)"

            st.write(f"**Pré-visualização** ({len(df_cat)} linhas detectadas):")
            st.dataframe(df_cat.head(5), use_container_width=True)

            st.markdown("**Mapeamento de colunas** — confirme ou ajuste:")
            c1, c2, c3, c4 = st.columns(4)
            col_cod    = c1.selectbox("Código *",    colunas, index=colunas.index(_sugerir(["cód","cod","codigo","ref"])))
            col_nome   = c2.selectbox("Produto *",   colunas, index=colunas.index(_sugerir(["produto","nome","descrição","descricao"])))
            col_fam    = c3.selectbox("Família",     colunas, index=colunas.index(_sugerir(["famil","categoria","grupo","linha"])))
            col_peso   = c4.selectbox("Peso",        colunas, index=colunas.index(_sugerir(["peso","gramatura","kg","g"])))
            c5, c6, c7 = st.columns(3)
            col_qtd    = c5.selectbox("Qtde CX",     colunas, index=colunas.index(_sugerir(["qtd","quant","cx","caixa","pcs"])))
            col_und    = c6.selectbox("UND",         colunas, index=colunas.index(_sugerir(["und","unidade","emb","embalagem"])))
            col_preco  = c7.selectbox("Preço",       colunas, index=colunas.index(_sugerir(["preço","preco","valor","price"])))

            if col_cod == "(ignorar)" or col_nome == "(ignorar)":
                st.warning("Código e Produto são obrigatórios.")
            else:
                col_map = {
                    "codigo_fab": col_cod,
                    "nome":       col_nome,
                    "familia":    col_fam    if col_fam    != "(ignorar)" else None,
                    "peso":       col_peso   if col_peso   != "(ignorar)" else None,
                    "qtde_cx":    col_qtd    if col_qtd    != "(ignorar)" else None,
                    "und":        col_und    if col_und    != "(ignorar)" else None,
                    "preco":      col_preco  if col_preco  != "(ignorar)" else None,
                }
                if st.button("✅ Confirmar importação do catálogo", use_container_width=True):
                    n = banco.importar_catalogo_df(fab["id"], df_cat, col_map=col_map)
                    st.success(f"{n} produtos importados/atualizados para {fab_sel}.")
                    st.rerun()
        except Exception as e:
            st.error(f"Erro ao ler arquivo: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: CÓDIGOS DA REDE
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "🔢 Códigos da Rede":
    _c1, _c2 = st.columns([5, 1])
    with _c1:
        st.title("Códigos da Rede")
    with _c2:
        if st.button("🔄 Atualizar", key="ref_cod"):
            st.rerun()
    st.caption("Mapeamento entre o código da fábrica e o código que cada rede usa internamente.")

    import unicodedata as _ud

    def _norm(t):
        return _ud.normalize("NFD", str(t).lower()).encode("ascii","ignore").decode("ascii").strip()

    fabricas = banco.listar_fabricas()
    redes    = banco.listar_redes()
    nomes_fab  = [f["nome"] for f in fabricas]
    nomes_rede = [r["nome"] for r in redes]

    # ── Importar todas as fábricas de uma vez ─────────────────────────────────
    with st.expander("📥 Importar códigos de múltiplas fábricas de uma vez", expanded=False):
        st.caption("Envie uma planilha com colunas: Fábrica · Cód. Fábrica · Cód. Rede — uma linha por produto.")
        import io as _io_m

        rede_multi_sel = st.selectbox("Rede / Cliente", nomes_rede, key="multi_rede")
        rede_multi = next((r for r in redes if r["nome"] == rede_multi_sel), None)

        arq_multi = st.file_uploader("Planilha (.xlsx)", type=["xlsx","xls"], key="upload_multi")
        if arq_multi is not None:
            _b = arq_multi.read()
            if _b:
                st.session_state["_cache_multi"] = (arq_multi.name, _b)

        _multi_cache = st.session_state.get("_cache_multi")
        if _multi_cache and rede_multi:
            _mn, _mb = _multi_cache
            try:
                df_m = pd.read_excel(_io_m.BytesIO(_mb), dtype=str)
                df_m.columns = [str(c).strip() for c in df_m.columns]
                df_m = df_m.dropna(how="all")
                cols_m = ["(ignorar)"] + list(df_m.columns)

                def _sg(kws):
                    for c in df_m.columns:
                        if any(k in c.lower() for k in kws): return c
                    return "(ignorar)"

                st.dataframe(df_m.head(4), use_container_width=True, hide_index=True)
                mc1, mc2, mc3 = st.columns(3)
                col_mfab  = mc1.selectbox("Coluna — Fábrica *",      cols_m, index=cols_m.index(_sg(["fabrica","fábrica","fab","fornec"])), key="m_fab")
                col_mcfab = mc2.selectbox("Coluna — Cód. Fábrica *", cols_m, index=cols_m.index(_sg(["cod fab","cód fab","codigo fab","código fab","cod. fab"])), key="m_cfab")
                col_mcred = mc3.selectbox("Coluna — Cód. Rede *",    cols_m, index=cols_m.index(_sg(["cod rede","cód rede","codigo rede","código rede","cod. rede","rede","cliente"])), key="m_cred")

                if col_mfab != "(ignorar)" and col_mcfab != "(ignorar)" and col_mcred != "(ignorar)":
                    df_m2 = df_m[[col_mfab, col_mcfab, col_mcred]].copy()
                    df_m2.columns = ["fabrica_nome","codigo_fab","codigo_rede"]
                    df_m2 = df_m2.dropna(subset=["codigo_fab","codigo_rede"])
                    df_m2 = df_m2[~df_m2["codigo_fab"].isin(["nan","None",""])]
                    df_m2 = df_m2[~df_m2["codigo_rede"].isin(["nan","None",""])]

                    def _match_fab(nome_planilha):
                        n = _norm(nome_planilha)
                        for f in fabricas:
                            if _norm(f["nome"]) in n or n in _norm(f["nome"]):
                                return f
                            if _norm(f["nome_faturamento"]) in n or n in _norm(f["nome_faturamento"]):
                                return f
                        return None

                    def _match_fab_id(nome_planilha):
                        f = _match_fab(nome_planilha)
                        return f["id"] if f else None

                    df_m2["fab_id"] = df_m2["fabrica_nome"].apply(_match_fab_id)
                    nao_rec = df_m2[df_m2["fab_id"].isna()]["fabrica_nome"].unique().tolist()
                    df_ok   = df_m2[df_m2["fab_id"].notna()].copy()

                    resumo = df_ok.groupby("fabrica_nome").size().reset_index(name="qtd")
                    st.markdown(f"**{len(df_ok)} pares válidos** encontrados:")
                    st.dataframe(resumo.rename(columns={"fabrica_nome":"Fábrica","qtd":"Qtd"}),
                                 use_container_width=True, hide_index=True)
                    if nao_rec:
                        st.warning(f"Fábricas não reconhecidas: {', '.join(nao_rec)} — verifique o nome exato cadastrado.")

                    if len(df_ok) > 0 and st.button("✅ Importar todas as fábricas", use_container_width=True, type="primary", key="btn_multi"):
                        total_ok = total_nao = 0
                        for fab_id, grupo in df_ok.groupby("fab_id"):
                            fab_obj = next((f for f in fabricas if f["id"] == fab_id), None)
                            if not fab_obj:
                                continue
                            df_par = grupo[["codigo_fab","codigo_rede"]].copy()
                            n_imp, nao_enc = banco.importar_codigos_rede_df(fab_obj["id"], rede_multi["id"], df_par)
                            total_ok  += n_imp
                            total_nao += len(nao_enc)
                        st.success(f"✅ {total_ok} códigos importados.")
                        if total_nao:
                            st.warning(f"{total_nao} produto(s) não encontrados no catálogo — importe o catálogo primeiro.")
                        st.rerun()
            except Exception as _e:
                st.error(f"Erro: {_e}")

    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        fab_sel  = st.selectbox("Fábrica", nomes_fab, key="cr_fab")
    with col2:
        rede_sel = st.selectbox("Rede", nomes_rede, key="cr_rede")

    fab  = next((f for f in fabricas if f["nome"] == fab_sel), None)
    rede = next((r for r in redes if r["nome"] == rede_sel), None)

    if fab and rede:
        codigos = banco.listar_codigos_rede(fab["id"], rede["id"])
        if codigos:
            df = pd.DataFrame([{
                "Cód. Fab":    c["codigo_fab"],
                "Produto":     c["produto_nome"],
                "Cód. Rede":   c["codigo_rede"],
                "Ativo":       "✅" if c["ativo"] else "❌",
            } for c in codigos])
            st.dataframe(df, use_container_width=True, hide_index=True, height=300)
            st.caption(f"{len(codigos)} produtos mapeados para {fab_sel} × {rede_sel}")
        else:
            st.info(f"Nenhum código cadastrado para {fab_sel} × {rede_sel}.")

        st.markdown("---")
        st.markdown("**Importar mapeamento (.xlsx)**")
        st.caption("Envie qualquer planilha — o sistema detecta as colunas automaticamente e você confirma o mapeamento.")

        import io as _io

        arq_cod = st.file_uploader("Selecione a planilha de mapeamento",
                                   type=["xlsx", "xls"], key="upload_codigos")
        # Atualiza cache somente quando há bytes reais (evita esvaziar no rerun)
        if arq_cod is not None:
            _bytes_cod = arq_cod.read()
            if _bytes_cod:
                st.session_state["_cache_codigos"] = (arq_cod.name, _bytes_cod)

        _cod_cache = st.session_state.get("_cache_codigos")

        if _cod_cache:
            _cod_name, _cod_bytes = _cod_cache

            # Botão de reimportação rápida (se já houver mapeamento salvo)
            _map_salvo = st.session_state.get("_map_codigos")
            if _map_salvo and arq_cod is None:
                st.info(f"📄 **{_cod_name}** em cache  |  Mapeamento salvo: {_map_salvo['fab']} → {_map_salvo['rede']}")
                if st.button("🔁 Re-importar com mapeamento salvo", use_container_width=True, key="reimp_cod"):
                    try:
                        df_ri = pd.read_excel(_io.BytesIO(_cod_bytes), header=None)
                        # detectar cabeçalho
                        hr = 0
                        for ii, rr in df_ri.iterrows():
                            if any(kw in " ".join(str(v).lower() for v in rr.values)
                                   for kw in ["cód","cod","código","produto","nome"]):
                                hr = ii; break
                        df_ri.columns = [str(c).strip() for c in df_ri.iloc[hr]]
                        df_ri = df_ri.iloc[hr+1:].reset_index(drop=True).dropna(how="all")
                        df_pi = pd.DataFrame({
                            "codigo_fab":  df_ri[_map_salvo["fab"]].astype(str).str.strip(),
                            "codigo_rede": df_ri[_map_salvo["rede"]].astype(str).str.strip(),
                        })
                        df_pi = df_pi[(df_pi["codigo_fab"] != "") & (df_pi["codigo_rede"] != "") &
                                      (~df_pi["codigo_fab"].isin(["nan","None"])) &
                                      (~df_pi["codigo_rede"].isin(["nan","None"]))]
                        n, nao_enc = banco.importar_codigos_rede_df(fab["id"], rede["id"], df_pi)
                        if n > 0:
                            st.success(f"✅ {n} códigos reimportados.")
                        if nao_enc:
                            st.warning(f"{len(nao_enc)} produto(s) não encontrado(s): {', '.join(nao_enc[:10])}")
                        st.rerun()
                    except Exception as _e:
                        st.error(f"Erro: {_e}")

            try:
                df_raw_cod = pd.read_excel(_io.BytesIO(_cod_bytes), header=None)
                # Detectar cabeçalho
                header_row = 0
                for i, row in df_raw_cod.iterrows():
                    vals = [str(v).lower() for v in row.values]
                    joined = " ".join(vals)
                    if any(kw in joined for kw in ["cód", "cod", "código", "produto", "nome"]):
                        header_row = i
                        break
                df_cod = df_raw_cod.copy()
                # Deduplica nomes de colunas
                raw_cols_cod = [str(c).strip() for c in df_raw_cod.iloc[header_row]]
                seen_cod = {}
                dedup_cols_cod = []
                for c in raw_cols_cod:
                    if c in seen_cod:
                        seen_cod[c] += 1
                        dedup_cols_cod.append(f"{c}_{seen_cod[c]}" if c != "nan" else f"_col{seen_cod[c]}")
                    else:
                        seen_cod[c] = 0
                        dedup_cols_cod.append(c if c != "nan" else f"_col{len(dedup_cols_cod)}")
                df_cod.columns = dedup_cols_cod
                df_cod = df_cod.iloc[header_row + 1:].reset_index(drop=True)
                df_cod = df_cod.dropna(how="all")
                df_cod = df_cod[[c for c in df_cod.columns if not c.startswith("_col")]]

                colunas_cod = ["(ignorar)"] + list(df_cod.columns)

                def _sug_cod(keywords):
                    for col in df_cod.columns:
                        c = str(col).lower()
                        if any(kw in c for kw in keywords):
                            return col
                    return "(ignorar)"

                with st.expander(f"📋 Mapeamento de colunas — {_cod_name}", expanded=(arq_cod is not None)):
                    st.write(f"**Pré-visualização** ({len(df_cod)} linhas detectadas):")
                    st.dataframe(df_cod.head(5), use_container_width=True)

                    st.markdown("**Confirme ou ajuste as colunas:**")
                    cc1, cc2 = st.columns(2)
                    col_cf = cc1.selectbox(
                        "Código da Fábrica *",
                        colunas_cod,
                        index=colunas_cod.index(_sug_cod(["fab", "fornec", "quata", "allfood", "prieto", "bufa", "sao vic", "villa", "cod fab", "codigo fab"])),
                        key="map_cod_fab"
                    )
                    col_cr = cc2.selectbox(
                        f"Código da Rede ({rede_sel}) *",
                        colunas_cod,
                        index=colunas_cod.index(_sug_cod(["rede", "cliente", "assai", "atacad", "gbarbosa", "carrefour", "extra", "bomprec", "soberano", "cod rede"])),
                        key="map_cod_rede"
                    )

                    if col_cf == "(ignorar)" or col_cr == "(ignorar)":
                        st.warning("Selecione as duas colunas obrigatórias.")
                    else:
                        df_para_import = pd.DataFrame({
                            "codigo_fab":  df_cod[col_cf].astype(str).str.strip(),
                            "codigo_rede": df_cod[col_cr].astype(str).str.strip(),
                        })
                        df_para_import = df_para_import[
                            (df_para_import["codigo_fab"] != "") &
                            (df_para_import["codigo_rede"] != "") &
                            (~df_para_import["codigo_fab"].isin(["nan", "None"])) &
                            (~df_para_import["codigo_rede"].isin(["nan", "None"]))
                        ]
                        st.caption(f"{len(df_para_import)} pares válidos encontrados.")
                        if st.button("✅ Confirmar importação dos códigos", use_container_width=True):
                            n, nao_enc = banco.importar_codigos_rede_df(fab["id"], rede["id"], df_para_import)
                            # Salva mapeamento para re-importação futura
                            st.session_state["_map_codigos"] = {"fab": col_cf, "rede": col_cr}
                            if n > 0:
                                st.success(f"{n} códigos importados/atualizados.")
                            if nao_enc:
                                st.warning(
                                    f"{len(nao_enc)} produto(s) não encontrado(s) no catálogo de {fab_sel} "
                                    f"— importe o catálogo primeiro na aba **Produtos**.\n\n"
                                    f"Códigos não encontrados: {', '.join(nao_enc[:10])}"
                                    + (" ..." if len(nao_enc) > 10 else "")
                                )
                            if n == 0 and not nao_enc:
                                st.info("Nenhum par código-produto encontrado. Verifique o mapeamento de colunas.")
                            st.rerun()
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")

        st.markdown("---")
        with st.expander("➕ Adicionar código individual"):
            prods = banco.listar_produtos(fab["id"])
            if prods:
                opcoes = {f"{p['codigo_fab']} — {p['nome']}": p["id"] for p in prods}
                prod_sel_str = st.selectbox("Produto", list(opcoes.keys()))
                prod_id_sel  = opcoes[prod_sel_str]
                cod_in = st.text_input("Código da rede")
                if st.button("💾 Salvar código"):
                    if cod_in:
                        banco.salvar_codigo_rede(prod_id_sel, rede["id"], cod_in)
                        st.success("Salvo!")
                        st.rerun()
            else:
                st.info("Importe o catálogo da fábrica primeiro.")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: TABELA DE PREÇO
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "💰 Tabela de Preço":
    st.title("Tabela de Preço por Rede / Cliente")
    st.caption(
        "Importe a tabela de preços negociada com cada cliente. "
        "O sistema usará o preço e a unidade (PC, CX, KG) desta tabela ao gerar o relatório Excel."
    )

    fabricas = banco.listar_fabricas()
    redes    = banco.listar_redes()

    col1, col2 = st.columns(2)
    with col1:
        fab_sel  = st.selectbox("Fábrica", [f["nome"] for f in fabricas], key="tp_fab")
    with col2:
        rede_sel = st.selectbox("Rede / Cliente", [r["nome"] for r in redes], key="tp_rede")

    fab_obj  = next((f for f in fabricas if f["nome"] == fab_sel), None)
    rede_obj = next((r for r in redes    if r["nome"] == rede_sel), None)

    if fab_obj and rede_obj:
        tab_atual = banco.obter_tabela_precos(rede_obj["id"], fab_obj["id"])

        # ── Tabela atual ──────────────────────────────────────────────────────
        if tab_atual:
            st.markdown(f"**{len(tab_atual)} itens cadastrados** — {rede_sel} × {fab_sel}")
            df_atual = pd.DataFrame([{
                "Cód. Fab":   r["codigo_fab"],
                "Produto":    r["produto_nome"] or r["descricao"] or "—",
                "Unidade":    r["unidade"] or "—",
                "Preço (R$)": f"R$ {float(r['preco']):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                "Vinculado":  "✅" if r["produto_nome"] else "⚠️ não encontrado",
            } for r in tab_atual])
            st.dataframe(df_atual, use_container_width=True, hide_index=True, height=280)

            nao_vinc = sum(1 for r in tab_atual if not r["produto_nome"])
            if nao_vinc:
                st.warning(
                    f"⚠️ {nao_vinc} produto(s) não vinculados — verifique se o Cód. Fab. "
                    "existe no cadastro de Produtos desta fábrica."
                )
        else:
            st.info("Nenhuma tabela de preços cadastrada para esta combinação. Faça o upload abaixo.")

        st.markdown("---")

        # ── Upload ────────────────────────────────────────────────────────────
        st.subheader("📂 Importar tabela de preços (.xlsx)")
        st.markdown(
            "O arquivo Excel deve ter **pelo menos** as colunas de: "
            "**código da fábrica**, **unidade** e **preço**. "
            "Após o upload, mapeie as colunas correspondentes abaixo."
        )

        arquivo = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls"], key="tp_upload")

        if arquivo:
            try:
                df_raw = pd.read_excel(arquivo, dtype=str)
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                st.success(f"Arquivo lido: {len(df_raw)} linhas, colunas: {list(df_raw.columns)}")

                st.markdown("**Mapeamento de colunas:**")
                cols_disponiveis = ["(não usar)"] + list(df_raw.columns)

                m1, m2, m3, m4 = st.columns(4)
                with m1:
                    col_cod  = st.selectbox("Cód. Fab (obrigatório)", cols_disponiveis,
                                            index=next((i+1 for i, c in enumerate(df_raw.columns)
                                                        if any(x in c.lower() for x in ["cod","cód","código","codigo","ref"])), 0),
                                            key="tp_col_cod")
                with m2:
                    col_desc = st.selectbox("Descrição / Nome", cols_disponiveis,
                                            index=next((i+1 for i, c in enumerate(df_raw.columns)
                                                        if any(x in c.lower() for x in ["desc","nome","produto","item"])), 0),
                                            key="tp_col_desc")
                with m3:
                    col_und  = st.selectbox("Unidade (PC/CX/KG)", cols_disponiveis,
                                            index=next((i+1 for i, c in enumerate(df_raw.columns)
                                                        if any(x in c.lower() for x in ["und","unid","unit"])), 0),
                                            key="tp_col_und")
                with m4:
                    col_prec = st.selectbox("Preço", cols_disponiveis,
                                            index=next((i+1 for i, c in enumerate(df_raw.columns)
                                                        if any(x in c.lower() for x in ["prec","preço","valor","price"])), 0),
                                            key="tp_col_prec")

                if col_cod != "(não usar)":
                    # Preview
                    preview_cols = {col_cod: "Cód. Fab"}
                    if col_desc != "(não usar)": preview_cols[col_desc] = "Descrição"
                    if col_und  != "(não usar)": preview_cols[col_und]  = "Unidade"
                    if col_prec != "(não usar)": preview_cols[col_prec] = "Preço"

                    df_prev = df_raw[list(preview_cols.keys())].rename(columns=preview_cols)
                    df_prev = df_prev.dropna(subset=["Cód. Fab"]).head(10)
                    st.markdown("**Prévia (10 primeiras linhas):**")
                    st.dataframe(df_prev, use_container_width=True, hide_index=True)

                    st.markdown("")
                    if st.button("💾 Salvar tabela de preços", type="primary", use_container_width=True, key="tp_salvar"):
                        if col_prec == "(não usar)":
                            st.error("Selecione a coluna de Preço.")
                        else:
                            # Monta lista de itens
                            items = []
                            for _, row in df_raw.iterrows():
                                cod = str(row.get(col_cod, "") or "").strip()
                                if not cod or cod.lower() in ("nan", "none", ""):
                                    continue
                                preco_raw = str(row.get(col_prec, "0") or "0").strip()
                                preco_raw = preco_raw.replace("R$","").replace(" ","").replace(".","").replace(",",".").strip()
                                try:
                                    preco_val = float(preco_raw)
                                except ValueError:
                                    preco_val = 0.0
                                items.append({
                                    "codigo_fab": cod,
                                    "descricao":  str(row.get(col_desc, "") or "") if col_desc != "(não usar)" else "",
                                    "unidade":    str(row.get(col_und,  "") or "") if col_und  != "(não usar)" else "",
                                    "preco":      preco_val,
                                })

                            if items:
                                banco.salvar_tabela_precos(rede_obj["id"], fab_obj["id"], items)
                                vinc = sum(1 for it in items if it["codigo_fab"])
                                st.success(f"✅ {len(items)} itens salvos para {rede_sel} × {fab_sel}!")
                                st.rerun()
                            else:
                                st.error("Nenhum item válido encontrado. Verifique o mapeamento.")
                else:
                    st.warning("Selecione a coluna com o Código da Fábrica.")

            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")

        # ── Limpar tabela ─────────────────────────────────────────────────────
        if tab_atual:
            st.markdown("---")
            with st.expander("🗑️ Remover tabela de preços"):
                st.warning(f"Isso removerá todos os {len(tab_atual)} itens da tabela de {rede_sel} × {fab_sel}.")
                if st.button("🗑️ Confirmar remoção", key="tp_limpar"):
                    banco.limpar_tabela_precos(rede_obj["id"], fab_obj["id"])
                    st.success("Tabela removida.")
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: IMPORTAR FATURAMENTO
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "📥 Importar Faturamento":
    _c1, _c2 = st.columns([5, 1])
    with _c1:
        st.title("Importar Faturamento")
    with _c2:
        if st.button("🔄 Atualizar", key="ref_fat"):
            st.rerun()
    st.caption("Selecione um ou mais arquivos. O período é detectado pelo nome do arquivo automaticamente.")

    MESES = {
        1: "Janeiro", 2: "Fevereiro", 3: "Marco", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    MESES_PT = {
        "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
        "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
        "janeiro": 1, "fevereiro": 2, "marco": 3, "março": 3, "abril": 4,
        "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
        "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
    }

    def _detectar_periodo(nome_arquivo):
        """Tenta extrair YYYY-MM do nome do arquivo."""
        import re
        n = nome_arquivo.lower().replace("_", "-").replace(" ", "-")
        # Padrão YYYY-MM ou MM-YYYY
        m = re.search(r'(20\d{2})[^\d]?(\d{2})', n)
        if m:
            ano, mes = int(m.group(1)), int(m.group(2))
            if 1 <= mes <= 12:
                return f"{ano}-{mes:02d}", f"{MESES.get(mes, mes)}/{ano}"
        m = re.search(r'(\d{2})[^\d]?(20\d{2})', n)
        if m:
            mes, ano = int(m.group(1)), int(m.group(2))
            if 1 <= mes <= 12:
                return f"{ano}-{mes:02d}", f"{MESES.get(mes, mes)}/{ano}"
        # Nome do mês por extenso
        ano_atual = datetime.date.today().year
        for nome_mes, num_mes in MESES_PT.items():
            if nome_mes in n:
                m_ano = re.search(r'20\d{2}', n)
                ano = int(m_ano.group()) if m_ano else ano_atual
                return f"{ano}-{num_mes:02d}", f"{MESES.get(num_mes, num_mes)}/{ano}"
        return None, None

    historico = banco.listar_importacoes()

    arqs_fat = st.file_uploader(
        "Selecione os arquivos de faturamento (.xls ou .xlsx) — pode selecionar vários",
        type=["xls", "xlsx"],
        key="upload_fat",
        accept_multiple_files=True,
    )

    # Cache novos uploads; só atualiza se bytes não estiverem vazios
    if arqs_fat:
        _novos = [(f.name, f.size, f.read()) for f in arqs_fat]
        _novos = [t for t in _novos if t[2]]  # descarta uploads com bytes vazios (rerun)
        if _novos:
            st.session_state["_cache_fat"] = _novos

    fat_lista = st.session_state.get("_cache_fat", [])

    if fat_lista and not arqs_fat:
        st.info(f"📦 {len(fat_lista)} arquivo(s) em cache. Clique 🔄 Atualizar para reprocessar ou envie novos arquivos.")
        if st.button("🗑️ Limpar arquivos em cache", key="rm_fat"):
            del st.session_state["_cache_fat"]
            st.rerun()

    if fat_lista:
        st.markdown("**Arquivos detectados:**")
        ano_atual = datetime.date.today().year
        mes_atual = datetime.date.today().month

        configs = []
        for i, (arq_nome, arq_size, arq_bytes) in enumerate(fat_lista):
            tag_auto, label_auto = _detectar_periodo(arq_nome)
            with st.expander(f"📄 {arq_nome} ({arq_size/1024:.0f} KB)", expanded=True):
                if tag_auto:
                    st.success(f"Período detectado: **{label_auto}**")
                    usar_auto = st.checkbox("Usar período detectado", value=True, key=f"auto_{i}")
                else:
                    st.warning("Período não detectado no nome do arquivo — selecione manualmente:")
                    usar_auto = False

                if not tag_auto or not usar_auto:
                    cc1, cc2 = st.columns(2)
                    with cc1:
                        mes_m = cc1.selectbox("Mês", list(MESES.keys()),
                                              format_func=lambda x: MESES[x],
                                              index=mes_atual - 1, key=f"mes_{i}")
                    with cc2:
                        ano_m = cc2.selectbox("Ano", [ano_atual - 1, ano_atual],
                                              index=1, key=f"ano_{i}")
                    tag_final   = f"{ano_m}-{mes_m:02d}"
                    label_final = f"{MESES[mes_m]}/{ano_m}"
                else:
                    tag_final   = tag_auto
                    label_final = label_auto

                ja_importado = next((h for h in historico if h["arquivo"] == tag_final), None)
                if ja_importado:
                    st.warning(f"⚠️ {label_final} já importado ({ja_importado['n_linhas']} lançamentos). Re-importar substitui.")

                configs.append((arq_bytes, tag_final, label_final))

        if st.button("📥 Importar todos os arquivos", type="primary", use_container_width=True):
            for conteudo, tag, label in configs:
                with st.spinner(f"Processando {label}..."):
                    try:
                        stats = imp.importar_faturamento(conteudo, tag)
                        if stats["gravados"] > 0:
                            st.success(f"✅ **{label}** — {stats['gravados']} lançamentos gravados "
                                       f"(sem produto: {stats['sem_produto']})")
                        else:
                            st.warning(f"⚠️ **{label}** — nenhum lançamento gravado. "
                                       f"Sem produto: {stats['sem_produto']}. Verifique os catálogos.")
                    except Exception as e:
                        st.error(f"❌ **{label}** — Erro: {e}")
            st.rerun()

    st.markdown("---")
    st.markdown("**Histórico de importações por mês**")
    historico = banco.listar_importacoes()
    if historico:
        rows_h = []
        for h in historico:
            # Tenta montar label amigável a partir do tag AAAA-MM
            try:
                partes = h["arquivo"].split("-")
                lbl = f"{MESES[int(partes[1])]}/{partes[0]}"
            except Exception:
                lbl = h["arquivo"]
            rows_h.append({
                "Periodo":         lbl,
                "Data importacao": h["data_importacao"][:16],
                "Lancamentos":     h["n_linhas"],
                "Status":          h["status"],
            })
        st.dataframe(pd.DataFrame(rows_h), use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma importacao realizada ainda.")

    with st.expander("🔍 Diagnóstico — o que está no banco"):
        diag = banco.diagnostico_faturamento()
        if diag:
            st.dataframe(pd.DataFrame(diag), use_container_width=True, hide_index=True)
        else:
            st.warning("Nenhum registro de faturamento encontrado no banco — verifique se a importação está gravando.")

    with st.expander("Limpar todos os dados de faturamento"):
        st.warning("Isso apaga TODOS os lancamentos de faturamento do banco. Os catalogos e codigos nao sao afetados.")
        confirm = st.text_input("Digite CONFIRMAR para apagar:")
        if confirm == "CONFIRMAR":
            if st.button("Apagar faturamento", type="primary"):
                banco.limpar_todo_faturamento()
                st.success("Faturamento apagado.")
                st.rerun()

# ── PÁGINA: USUÁRIOS (admin only) ─────────────────────────────────────────────
elif pagina == "👥 Usuários" and _is_admin:
    st.title("👥 Gestão de Usuários")
    st.caption("Usuários comuns têm acesso apenas à página **Gerar Relatório**.")

    # --- Cadastrar novo usuário ---
    with st.expander("➕ Cadastrar novo usuário", expanded=False):
        with st.form("form_novo_usuario"):
            nu_nome     = st.text_input("Nome completo")
            nu_username = st.text_input("Usuário (login)")
            nu_senha    = st.text_input("Senha", type="password")
            nu_senha2   = st.text_input("Confirmar senha", type="password")
            salvar_u    = st.form_submit_button("Cadastrar", use_container_width=True, type="primary")
            if salvar_u:
                if not nu_nome or not nu_username or not nu_senha:
                    st.error("Preencha todos os campos.")
                elif nu_senha != nu_senha2:
                    st.error("As senhas não conferem.")
                else:
                    ok, msg = banco.criar_usuario(nu_username, nu_nome, nu_senha)
                    if ok:
                        st.success(f"Usuário **{nu_username}** cadastrado com sucesso.")
                        st.rerun()
                    else:
                        st.error(f"Erro: {msg}")

    # --- Lista de usuários ---
    st.markdown("### Usuários cadastrados")
    usuarios = banco.listar_usuarios()
    if not usuarios:
        st.info("Nenhum usuário cadastrado ainda.")
    else:
        for u in usuarios:
            col_n, col_u, col_ac, col_del, col_pwd = st.columns([2, 1.5, 1.8, 1, 1])
            col_n.markdown(f"**{u['nome']}**")
            col_u.markdown(f"`{u['username']}`")
            ultimo = u.get("ultimo_acesso")
            col_ac.caption(f"Último acesso: {str(ultimo)[:16] if ultimo else 'nunca'}")
            if col_del.button("🗑️", key=f"del_{u['id']}", help="Excluir usuário"):
                banco.deletar_usuario(u["id"])
                st.success(f"Usuário **{u['username']}** excluído.")
                st.rerun()
            if col_pwd.button("🔑", key=f"pwd_{u['id']}", help="Alterar senha"):
                st.session_state[f"_trocar_senha_{u['id']}"] = True

            # Formulário inline de troca de senha
            if st.session_state.get(f"_trocar_senha_{u['id']}"):
                with st.form(f"form_pwd_{u['id']}"):
                    np1 = st.text_input("Nova senha", type="password", key=f"np1_{u['id']}")
                    np2 = st.text_input("Confirmar", type="password", key=f"np2_{u['id']}")
                    salvar_p = st.form_submit_button("Salvar", use_container_width=True)
                    if salvar_p:
                        if np1 != np2:
                            st.error("As senhas não conferem.")
                        elif not np1:
                            st.error("Senha não pode ser vazia.")
                        else:
                            banco.alterar_senha(u["id"], np1)
                            del st.session_state[f"_trocar_senha_{u['id']}"]
                            st.success("Senha alterada.")
                            st.rerun()

    st.markdown("---")
    st.caption("🔐 A senha do administrador é configurada nas secrets do Streamlit Cloud "
               "(chave `[auth]` → `admin_username` / `admin_password`).")
