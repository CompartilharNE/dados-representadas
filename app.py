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
                try:
                    banco.deletar_rede(r_edit["id"])
                    st.success("Excluído.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao excluir: {e}")


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
    st.caption("Envie qualquer planilha — o sistema detecta as colunas automaticamente e você confirma o mapeamento.")

    arq = st.file_uploader("Selecione o arquivo de catálogo", type=["xlsx", "xls"],
                            key="upload_catalogo")
    if arq and fab:
        try:
            df_raw = pd.read_excel(arq, header=None)

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
        st.caption("Envie qualquer planilha — o sistema detecta as colunas automaticamente e você confirma o mapeamento.")
        arq_cod = st.file_uploader("Selecione a planilha de mapeamento",
                                   type=["xlsx", "xls"], key="upload_codigos")
        if arq_cod:
            try:
                df_raw_cod = pd.read_excel(arq_cod, header=None)
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

                st.write(f"**Pré-visualização** ({len(df_cod)} linhas detectadas):")
                st.dataframe(df_cod.head(5), use_container_width=True)

                st.markdown("**Mapeamento de colunas** — confirme ou ajuste:")
                cc1, cc2 = st.columns(2)
                col_cf = cc1.selectbox(
                    f"Código da Fábrica *",
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
                    # Montar df padronizado para importação
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
# PÁGINA: IMPORTAR FATURAMENTO
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "📥 Importar Faturamento":
    st.title("Importar Faturamento")
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

    if arqs_fat:
        st.markdown("**Arquivos detectados:**")
        ano_atual = datetime.date.today().year
        mes_atual = datetime.date.today().month

        configs = []
        for i, arq in enumerate(arqs_fat):
            tag_auto, label_auto = _detectar_periodo(arq.name)
            with st.expander(f"📄 {arq.name} ({arq.size/1024:.0f} KB)", expanded=True):
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

                configs.append((arq, tag_final, label_final))

        if st.button("📥 Importar todos os arquivos", type="primary", use_container_width=True):
            for arq, tag, label in configs:
                with st.spinner(f"Processando {label}..."):
                    conteudo = arq.read()
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

    with st.expander("Limpar todos os dados de faturamento"):
        st.warning("Isso apaga TODOS os lancamentos de faturamento do banco. Os catalogos e codigos nao sao afetados.")
        confirm = st.text_input("Digite CONFIRMAR para apagar:")
        if confirm == "CONFIRMAR":
            if st.button("Apagar faturamento", type="primary"):
                banco.limpar_todo_faturamento()
                st.success("Faturamento apagado.")
                st.rerun()
