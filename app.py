"""
app.py — Interface Streamlit
Compartilhar NE — Dados Representadas
Rodar: streamlit run app.py --server.port 8501
"""
import streamlit as st
import pandas as pd
import datetime
import banco
import gerar_excel
import importar as imp

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
    [data-testid="stSidebar"] { background-color: #1e2130; }
    [data-testid="stSidebar"] * { color: rgba(255,255,255,0.75) !important; }
    [data-testid="stSidebar"] .sidebar-title { color: #ff4b4b !important; font-size: 1.1rem; font-weight: 600; }
    .metric-box { background: #f8f9fa; border-radius: 8px; padding: 12px 16px; text-align: center; }
    .metric-label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-val   { font-size: 26px; font-weight: 600; margin-top: 2px; }
    .green { color: #2e7d32; }
    .red   { color: #c62828; }
    .blue  { color: #1565c0; }
    .section-header { font-size: 13px; font-weight: 600; color: #444; margin: 8px 0 4px; }
    div[data-testid="stDownloadButton"] button { background-color: #1F4E79; color: white; font-weight: 600; width: 100%; }
</style>
""", unsafe_allow_html=True)

# ── NAVEGAÇÃO SIDEBAR ─────────────────────────────────────────────────────────
with st.sidebar:
    st.image("compartilhar_logo.svg", use_container_width=True)
    st.markdown('<div class="sidebar-title" style="margin-top:6px">📊 Dados Representadas</div>', unsafe_allow_html=True)
    st.markdown('<div style="font-size:11px;color:rgba(255,255,255,0.4);margin-bottom:12px">Compartilhar NE</div>', unsafe_allow_html=True)

    pagina = st.radio(
        "Navegação",
        ["🗂️ Gerar Relatório", "🏭 Fábricas", "🏪 Redes / Clientes",
         "📦 Produtos", "🔢 Códigos da Rede", "📥 Importar Faturamento"],
        label_visibility="collapsed",
    )
    st.markdown("---")
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
                    f"DadosRepresentadas_{fab_sel.replace('/', '-')}_"
                    f"{'-'.join(r.replace('/', '-').replace(' ', '') for r in redes_sel)}_"
                    f"{d_ini.strftime('%Y%m%d')}_{d_fim.strftime('%Y%m%d')}.xlsx"
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
    st.title("Fábricas")
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

        c1, c2 = st.columns([3, 1])
        with c1:
            if st.button("💾 Salvar", use_container_width=True):
                if nome_in and nome_fat_in:
                    banco.salvar_fabrica(nome_in, nome_fat_in, contato_in,
                                         fab_id=fab_edit["id"] if fab_edit else None)
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
    st.title("Redes / Clientes")
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

        c1, c2 = st.columns([3, 1])
        with c1:
            if st.button("💾 Salvar rede", use_container_width=True):
                if nome_in and filtro_in:
                    banco.salvar_rede(nome_in, filtro_in, estados_in, excluir_in,
                                      rede_id=r_edit["id"] if r_edit else None)
                    st.success("Salvo!")
                    st.rerun()
                else:
                    st.warning("Preencha nome e filtro.")
        with c2:
            if r_edit and st.button("🗑️ Excluir rede", use_container_width=True):
                banco.deletar_rede(r_edit["id"])
                st.success("Excluído.")
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: PRODUTOS
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "📦 Produtos":
    st.title("Produtos")
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
                "Preço":     p["preco"] or "—",
            } for p in prods])
            st.dataframe(df, use_container_width=True, hide_index=True, height=350)
            st.caption(f"{len(prods)} produtos cadastrados para {fab_sel}")
        else:
            st.info(f"Nenhum produto cadastrado para {fab_sel}. Importe o catálogo abaixo.")

    st.markdown("---")
    st.markdown("**Importar catálogo (.xlsx)**")
    st.caption("O arquivo deve ter colunas: Código (ou Cód), Produto (ou Nome), Família, Peso, Qtde CX, UND, Preço")

    arq = st.file_uploader("Selecione o arquivo de catálogo", type=["xlsx", "xls"],
                            key="upload_catalogo")
    if arq and fab:
        try:
            df_cat = pd.read_excel(arq, header=None)
            # Auto-detectar linha de cabeçalho
            header_row = 0
            for i, row in df_cat.iterrows():
                vals = [str(v).lower() for v in row.values]
                if any(kw in " ".join(vals) for kw in ["código", "cod", "produto", "nome"]):
                    header_row = i
                    break
            df_cat.columns = df_cat.iloc[header_row]
            df_cat = df_cat.iloc[header_row + 1:].reset_index(drop=True)
            df_cat = df_cat.dropna(how="all")

            st.write(f"**Pré-visualização** ({len(df_cat)} linhas):")
            st.dataframe(df_cat.head(10), use_container_width=True)

            if st.button("✅ Confirmar importação do catálogo", use_container_width=True):
                n = banco.importar_catalogo_df(fab["id"], df_cat)
                st.success(f"{n} produtos importados/atualizados para {fab_sel}.")
                st.rerun()
        except Exception as e:
            st.error(f"Erro ao ler arquivo: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: CÓDIGOS DA REDE
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "🔢 Códigos da Rede":
    st.title("Códigos da Rede")
    st.caption("Mapeamento entre o código da fábrica e o código que cada rede usa internamente.")

    fabricas = banco.listar_fabricas()
    redes    = banco.listar_redes()
    nomes_fab  = [f["nome"] for f in fabricas]
    nomes_rede = [r["nome"] for r in redes]

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
        st.caption("O arquivo deve ter colunas: Cód Fab (ou código da fábrica) e Cód Rede (ou código do cliente).")
        arq_cod = st.file_uploader("Selecione a planilha de mapeamento",
                                   type=["xlsx", "xls"], key="upload_codigos")
        if arq_cod:
            try:
                df_cod = pd.read_excel(arq_cod, header=None)
                # Detectar cabeçalho
                header_row = 0
                for i, row in df_cod.iterrows():
                    vals = [str(v).lower() for v in row.values]
                    if any("cód" in v or "cod" in v for v in vals):
                        header_row = i
                        break
                df_cod.columns = [str(c).strip() for c in df_cod.iloc[header_row]]
                df_cod = df_cod.iloc[header_row + 1:].reset_index(drop=True)
                df_cod = df_cod.dropna(how="all")

                # Renomear colunas para padrão
                col_map = {}
                for col in df_cod.columns:
                    cl = str(col).lower()
                    if "fab" in cl or "quatá" in cl or "fornecedor" in cl:
                        col_map[col] = "codigo_fab"
                    elif "atacadão" in cl or "rede" in cl or "cliente" in cl or "cod" in cl:
                        if "codigo_fab" not in col_map.values():
                            col_map[col] = "codigo_fab"
                        else:
                            col_map[col] = "codigo_rede"
                df_cod = df_cod.rename(columns=col_map)

                st.write(f"**Pré-visualização** ({len(df_cod)} linhas):")
                st.dataframe(df_cod.head(10), use_container_width=True)

                if st.button("✅ Confirmar importação dos códigos", use_container_width=True):
                    n = banco.importar_codigos_rede_df(fab["id"], rede["id"], df_cod)
                    st.success(f"{n} códigos importados/atualizados.")
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
# PÁGINA: IMPORTAR FATURAMENTO
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "📥 Importar Faturamento":
    st.title("Importar Faturamento")
    st.caption("Importe o faturamento mês a mês. Re-importar um mês substitui os dados anteriores daquele mês.")

    MESES = {
        1: "Janeiro", 2: "Fevereiro", 3: "Marco", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }

    ano_atual = datetime.date.today().year
    mes_atual = datetime.date.today().month

    col_m, col_a = st.columns(2)
    with col_m:
        mes_sel = st.selectbox(
            "Mês de referência",
            options=list(MESES.keys()),
            format_func=lambda x: MESES[x],
            index=mes_atual - 1,
            key="mes_fat"
        )
    with col_a:
        ano_sel = st.selectbox(
            "Ano",
            options=[ano_atual - 1, ano_atual],
            index=1,
            key="ano_fat"
        )

    periodo_tag = f"{ano_sel}-{mes_sel:02d}"
    periodo_label = f"{MESES[mes_sel]}/{ano_sel}"

    # Verificar se já existe importação para esse mês
    historico = banco.listar_importacoes()
    importado = next((h for h in historico if h["arquivo"] == periodo_tag), None)

    if importado:
        st.warning(f"⚠️ {periodo_label} já foi importado ({importado['n_linhas']} lançamentos em {importado['data_importacao'][:16]}). Re-importar vai substituir esses dados.")

    st.markdown("---")
    arq_fat = st.file_uploader(
        f"Selecione o arquivo de faturamento de {periodo_label} (.xls ou .xlsx)",
        type=["xls", "xlsx"],
        key="upload_fat",
    )

    if arq_fat:
        st.info(f"Arquivo: **{arq_fat.name}** — {arq_fat.size / 1024:.0f} KB")
        if st.button(f"📥 Importar {periodo_label}", type="primary", use_container_width=True):
            with st.spinner(f"Processando faturamento de {periodo_label}..."):
                conteudo = arq_fat.read()
                try:
                    stats = imp.importar_faturamento(conteudo, periodo_tag)
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Linhas lidas", stats["total"])
                    with col2:
                        st.metric("Gravados", stats["gravados"])
                    with col3:
                        st.metric("Sem produto cadastrado", stats["sem_produto"])
                    with col4:
                        st.metric("Erros", stats["erros"])

                    if stats["gravados"] > 0:
                        st.success(f"✅ {periodo_label} importado com sucesso! {stats['gravados']} lançamentos gravados.")
                    else:
                        st.warning("Nenhum lançamento foi gravado. Verifique se os catálogos estão cadastrados.")

                    if stats["sem_produto"] > stats["gravados"]:
                        st.info("💡 Muitos produtos sem cadastro. Importe os catálogos das fábricas na aba Produtos.")
                except Exception as e:
                    st.error(f"Erro ao importar: {e}")

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

    with st.expander("Limpar todos os dados de faturamento"):
        st.warning("Isso apaga TODOS os lancamentos de faturamento do banco. Os catalogos e codigos nao sao afetados.")
        confirm = st.text_input("Digite CONFIRMAR para apagar:")
        if confirm == "CONFIRMAR":
            if st.button("Apagar faturamento", type="primary"):
                banco.limpar_todo_faturamento()
                st.success("Faturamento apagado.")
                st.rerun()
