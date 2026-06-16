"""
gerar_excel.py — Geração do relatório Excel formatado
Compartilhar NE — Dados Representadas
"""
import io
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import banco


# ── ESTILOS ──────────────────────────────────────────────────────────────────
COR_TITULO    = "1F4E79"
COR_ROXO      = "CCCCFF"
COR_LARANJA   = "FFCC99"
COR_CINZA     = "F2F2F2"
COR_BRANCO    = "FFFFFF"
COR_VERDE_F   = "E2EFDA"
COR_VERM_F    = "FFF0F0"
COR_AZUL_2    = "2E75B6"
COR_LARANJA_E = "7F3F00"

_borda = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)


def _fill(c):
    return PatternFill("solid", fgColor=c)


def _fnt(bold=False, color="000000", size=10, italic=False):
    return Font(name="Calibri", bold=bold, color=color, size=size, italic=italic)


def _cel(ws, row, col, val, bold=False, color="000000", size=10,
         bg=None, align="left", fmt=None, italic=False, wrap=False):
    c = ws.cell(row=row, column=col, value=val)
    c.font = _fnt(bold, color, size, italic)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if bg:
        c.fill = _fill(bg)
    c.border = _borda
    if fmt:
        c.number_format = fmt
    return c


def _merge(ws, row, c1, c2, val, bg, color="000000", bold=False, size=10,
           align="left", h=17, italic=False):
    ws.merge_cells(f"{get_column_letter(c1)}{row}:{get_column_letter(c2)}{row}")
    c = ws.cell(row, c1, val)
    c.font = _fnt(bold, color, size, italic)
    c.alignment = Alignment(horizontal=align, vertical="center")
    if bg:
        c.fill = _fill(bg)
    c.border = _borda
    ws.row_dimensions[row].height = h
    return c


# ── ABA INDIVIDUAL DA FÁBRICA ─────────────────────────────────────────────────

def _criar_aba(wb, titulo, forn_nome, rede_nome, estados, periodo,
               prods_comprados, prods_nao_comprados, usar_familia=True):
    """Cria uma aba no workbook com os produtos comprados e não comprados."""
    safe = titulo.replace("/", "-").replace("\\", "-").replace("?", "").replace("*", "") \
                 .replace("[", "").replace("]", "").replace(":", "")[:31]
    ws = wb.create_sheet(safe)
    ws.sheet_view.showGridLines = False

    # Cabeçalho azul
    ncols = 10
    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    c = ws["A1"]
    c.value = "DADOS REPRESENTADAS"
    c.font = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = _fill(COR_TITULO)
    ws.row_dimensions[1].height = 28

    for i, (lb, vl) in enumerate([
        ("REDE:", rede_nome),
        ("FORNECEDOR:", forn_nome),
        ("PERÍODO:", periodo),
        ("ESTADOS:", estados),
    ], 2):
        ws.cell(i, 1).value = lb
        ws.cell(i, 1).font = _fnt(bold=True)
        ws.cell(i, 1).fill = _fill(COR_CINZA)
        ws.cell(i, 1).border = _borda
        ws.merge_cells(f"B{i}:{get_column_letter(ncols)}{i}")
        ws.cell(i, 2).value = vl
        ws.cell(i, 2).font = _fnt(bold=True, color=COR_TITULO)
        ws.cell(i, 2).fill = _fill(COR_CINZA)
        ws.cell(i, 2).border = _borda
        ws.row_dimensions[i].height = 18

    ws.row_dimensions[6].height = 6

    # Cabeçalho de colunas
    hdrs = ["CÓD. FAB", "CÓD. REDE", "PRODUTOS", "QTDE P/CX",
            "PESO CX", "UND.", "PREÇO", "QTD FATURADA", "VALOR (R$)", "Nº LOJAS"]
    bgs  = [COR_ROXO] * 7 + [COR_TITULO] * 3
    cors = ["000080"] * 7 + ["FFFFFF"] * 3
    for j, (h, bg, cor) in enumerate(zip(hdrs, bgs, cors), 1):
        c = ws.cell(7, j, h)
        c.font = _fnt(bold=True, color=cor, size=9)
        c.fill = _fill(bg)
        c.border = _borda
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[7].height = 24

    linha = 8
    fam_atual = None

    # COMPRADOS
    for i, p in enumerate(prods_comprados):
        fam = p.get("familia", "")
        if usar_familia and fam != fam_atual:
            fam_atual = fam
            _merge(ws, linha, 1, ncols, f"  {fam}", COR_AZUL_2, "FFFFFF", True, 9, "left", 16)
            linha += 1

        bg = COR_BRANCO if i % 2 == 0 else COR_VERDE_F
        vals = [
            (p.get("codigo_fab", ""), "center"),
            (p.get("codigo_rede", ""), "center"),
            (p.get("nome", ""), "left"),
            (p.get("qtde_cx", ""), "center"),
            (p.get("peso", ""), "center"),
            (p.get("und", ""), "center"),
            (p.get("preco", ""), "right"),
        ]
        for j, (v, al) in enumerate(vals, 1):
            _cel(ws, linha, j, v, align=al, bg=bg)
        _cel(ws, linha, 8, p.get("qtd_total", 0), align="right", bg=bg, fmt="#,##0.0")
        _cel(ws, linha, 9, p.get("valor_total", 0), align="right", bg=bg, fmt="R$ #,##0.00")
        _cel(ws, linha, 10, p.get("n_lojas", 0), align="center", bg=bg)
        ws.row_dimensions[linha].height = 17
        linha += 1

    # Separador SEM VENDA
    n_nc = len(prods_nao_comprados)
    ws.merge_cells(f"A{linha}:{get_column_letter(ncols)}{linha}")
    c = ws.cell(linha, 1, f"PRODUTOS SEM VENDA  ({n_nc} itens — OPORTUNIDADE COMERCIAL)")
    c.font = _fnt(bold=True, color=COR_LARANJA_E, size=10)
    c.fill = _fill(COR_LARANJA)
    c.border = _borda
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[linha].height = 20
    linha += 1

    fam_atual = None
    for i, p in enumerate(prods_nao_comprados):
        fam = p.get("familia", "")
        if usar_familia and fam != fam_atual:
            fam_atual = fam
            _merge(ws, linha, 1, ncols, f"  {fam}", "FFE5CC", COR_LARANJA_E, True, 9, "left", 16)
            linha += 1

        bg = COR_BRANCO if i % 2 == 0 else COR_VERM_F
        vals = [
            (p.get("codigo_fab", ""), "center"),
            (p.get("codigo_rede", ""), "center"),
            (p.get("nome", ""), "left"),
            (p.get("qtde_cx", ""), "center"),
            (p.get("peso", ""), "center"),
            (p.get("und", ""), "center"),
            (p.get("preco", ""), "right"),
        ]
        for j, (v, al) in enumerate(vals, 1):
            _cel(ws, linha, j, v, align=al, bg=bg)
        ws.row_dimensions[linha].height = 17
        linha += 1

    # Rodapé
    n_comp = len(prods_comprados)
    n_tot = n_comp + n_nc
    porc = round(n_comp / n_tot * 100, 1) if n_tot else 0
    valor = sum(p.get("valor_total", 0) for p in prods_comprados)
    linha += 1
    ws.merge_cells(f"A{linha}:{get_column_letter(ncols)}{linha}")
    c = ws.cell(linha, 1,
                f"Itens cadastrados: {n_tot}  |  Comprados: {n_comp}  |  "
                f"Não comprados: {n_nc}  |  Positivação: {porc}%  |  "
                f"Valor: R$ {valor:,.2f}")
    c.font = _fnt(italic=True, color="595959", size=9)
    c.fill = _fill(COR_CINZA)
    c.border = _borda
    c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[linha].height = 16

    # Larguras de coluna
    for col, w in zip("ABCDEFGHIJ", [11, 16, 44, 14, 12, 10, 18, 13, 18, 9]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A8"

    return n_comp, n_nc, valor


# ── ABA RESUMO ────────────────────────────────────────────────────────────────

def _criar_resumo(wb, periodo, resultados):
    """
    resultados: lista de dicts com:
        fab_nome, rede_nome, estados, n_cat, n_comp, n_ncomp, valor, titulo_aba, lojas
    """
    ws = wb.active
    ws.title = "📊 Resumo"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.outlinePr.summaryBelow = False

    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value = "DADOS REPRESENTADAS — COMPARTILHAR NE"
    c.font = Font(name="Calibri", bold=True, size=16, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = _fill(COR_TITULO)
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:J2")
    c = ws["A2"]
    c.value = (f"Período: {periodo}   |   Somente produtos cadastrados na rede   |   "
               "Clique no + na lateral para ver as lojas")
    c.font = Font(name="Calibri", italic=True, color="595959", size=10)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = _fill(COR_CINZA)
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 8

    hdrs = ["Fábrica", "Rede", "Estados", "Cadastrados\nna Rede",
            "Comprados", "Não\nComprados", "Positivação", "Valor (R$)", "Aba"]
    for j, h in enumerate(hdrs, 1):
        c = ws.cell(4, j, h)
        c.font = _fnt(bold=True, color="FFFFFF", size=9)
        c.fill = _fill(COR_TITULO)
        c.border = _borda
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 28

    COR_REGS = {
        "PE": ("EBF3FB", "1A5276", "D6EAF8", "EBF3FB"),
        "BA": ("FEF9E7", "7D6608", "FCF3CF", "FEF9E7"),
        "_":  ("F2F2F2", "555555", "EBEBEB", "F2F2F2"),
    }

    linha = 5
    for res in resultados:
        reg_key = "PE" if "PE" in res["rede_nome"] else ("BA" if "BA" in res["rede_nome"] else "_")
        cor_bg, cor_det, cor_l1, cor_l2 = COR_REGS[reg_key]
        porc = round(res["n_comp"] / res["n_cat"] * 100, 1) if res["n_cat"] else 0

        dados = [
            (res["fab_nome"], "left"),
            (res["rede_nome"], "left"),
            (res["estados"], "left"),
            (res["n_cat"], "center"),
            (res["n_comp"], "center"),
            (res["n_ncomp"], "center"),
            (f"{porc}%", "center"),
            (res["valor"], "right"),
            (res["titulo_aba"], "left"),
        ]
        for j, (v, al) in enumerate(dados, 1):
            c_ = _cel(ws, linha, j, v, align=al, bg=cor_bg, bold=(j == 1))
            if j == 8:
                c_.number_format = "R$ #,##0.00"
            if j == 5:
                c_.font = _fnt(bold=True, color="375623")
            if j == 6 and res["n_comp"] == 0:
                c_.font = _fnt(bold=True, color="9C0006")
        ws.row_dimensions[linha].height = 20
        linha += 1

        # Linhas de loja (outline 1, colapsadas)
        for k, loja in enumerate(res.get("lojas", [])):
            bg_l = cor_l1 if k % 2 == 0 else cor_l2
            ws.merge_cells(f"A{linha}:C{linha}")
            c_ = ws.cell(linha, 1, f"    ↳ {loja['nome_faturamento']}")
            c_.font = _fnt(size=9, italic=True, color=cor_det)
            c_.border = _borda
            c_.alignment = Alignment(horizontal="left", vertical="center")
            c_.fill = _fill(bg_l)
            ws.cell(linha, 4).value = loja.get("estado", "")
            ws.cell(linha, 4).font = _fnt(size=9, bold=True, color=cor_det)
            ws.cell(linha, 4).border = _borda
            ws.cell(linha, 4).fill = _fill(bg_l)
            ws.cell(linha, 4).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(linha, 5).value = loja.get("n_produtos", 0)
            ws.cell(linha, 5).font = _fnt(size=9, color=cor_det)
            ws.cell(linha, 5).border = _borda
            ws.cell(linha, 5).fill = _fill(bg_l)
            ws.cell(linha, 5).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(linha, 8).value = loja.get("valor_total", 0)
            ws.cell(linha, 8).number_format = "R$ #,##0.00"
            ws.cell(linha, 8).font = _fnt(size=9, color=cor_det)
            ws.cell(linha, 8).border = _borda
            ws.cell(linha, 8).fill = _fill(bg_l)
            ws.cell(linha, 8).alignment = Alignment(horizontal="right", vertical="center")
            for j in [6, 7, 9]:
                ws.cell(linha, j).fill = _fill(bg_l)
                ws.cell(linha, j).border = _borda
            ws.row_dimensions[linha].height = 15
            ws.row_dimensions[linha].outline_level = 1
            ws.row_dimensions[linha].hidden = True
            linha += 1

    for col, w in zip("ABCDEFGHI", [22, 26, 16, 12, 11, 11, 11, 18, 14]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"


# ── FUNÇÃO PRINCIPAL ─────────────────────────────────────────────────────────

def gerar_relatorio(fab_nome, redes_selecionadas, data_inicio, data_fim):
    """
    fab_nome: str (ex: "Quatá")
    redes_selecionadas: lista de str (ex: ["Atacadão PE", "Atacadão BA"])
    data_inicio, data_fim: str "YYYY-MM-DD" ou ""

    Retorna bytes do arquivo Excel ou None se não houver dados.
    """
    fab = banco.obter_fabrica_por_nome(fab_nome)
    if not fab:
        return None, "Fábrica não encontrada"

    periodo = ""
    if data_inicio and data_fim:
        def fmt(d):
            parts = d.split("-")
            return f"{parts[2]}/{parts[1]}/{parts[0]}" if len(parts) == 3 else d
        periodo = f"{fmt(data_inicio)} a {fmt(data_fim)}"
    elif data_inicio:
        periodo = f"A partir de {data_inicio}"
    elif data_fim:
        periodo = f"Até {data_fim}"
    else:
        periodo = "Todo o período"

    wb = openpyxl.Workbook()
    resultados = []

    for rede_nome in redes_selecionadas:
        rede = banco.obter_rede_por_nome(rede_nome)
        if not rede:
            continue

        prods_cat = banco.produtos_com_codigos(fab["id"], rede["id"])
        if not prods_cat:
            continue

        vendas = banco.vendas_por_produto(fab["id"], rede["id"], data_inicio, data_fim)
        lojas  = banco.vendas_por_loja(fab["id"], rede["id"], data_inicio, data_fim)

        comprados = []
        nao_comprados = []
        for p in prods_cat:
            vd = vendas.get(p["id"])
            if vd and vd["valor_total"] > 0:
                comprados.append({**p, **vd})
            else:
                nao_comprados.append(p)

        estados = rede["estados"].replace(",", " / ")
        titulo_aba = f"{fab_nome[:12]} {rede_nome.split()[-1]}"
        n_comp, n_nc, valor = _criar_aba(
            wb, titulo_aba, fab_nome, rede_nome, estados, periodo,
            comprados, nao_comprados, usar_familia=True
        )

        resultados.append({
            "fab_nome":   fab_nome,
            "rede_nome":  rede_nome,
            "estados":    estados,
            "n_cat":      len(prods_cat),
            "n_comp":     n_comp,
            "n_ncomp":    n_nc,
            "valor":      valor,
            "titulo_aba": titulo_aba,
            "lojas":      lojas,
        })

    if not resultados:
        return None, "Nenhum produto com código cadastrado para as redes selecionadas"

    _criar_resumo(wb, periodo, resultados)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), None
