"""
importar.py — Importação do arquivo de faturamento
Compartilhar NE — Dados Representadas
Suporta .xls (HTML disfarçado) e .xlsx real
"""
import io
import os
import unicodedata
import pandas as pd
from bs4 import BeautifulSoup
import banco


def _sem_acento(texto):
    """Remove acentos e normaliza para comparação."""
    return unicodedata.normalize("NFD", str(texto).lower()).encode("ascii", "ignore").decode("ascii")

# ── MAPEAMENTO ESTADO POR KEYWORD DE LOJA ────────────────────────────────────
MAPA_ESTADO = {
    # ── Cidades/regiões PE ────────────────────────────────────────────────────
    "arruda": "PE", "camaragibe": "PE", "caruaru": "PE", "casa amarela": "PE",
    "garanhuns": "PE", "igarassu": "PE", "iputinga": "PE", "jaboatão": "PE",
    "jaboatao": "PE", "olinda": "PE", "palmares": "PE", "paulista": "PE",
    "petrolina": "PE", "recife": "PE", "santa cruz do capibaribe": "PE",
    "serra talhada": "PE", "são lourenço da mata": "PE", "sao lourenco da mata": "PE",
    "vitória de santo antão": "PE", "vitoria de santo antao": "PE",
    "belo jardim": "PE", "candeias": "PE",
    # Bairros de Recife/Grande Recife — PE
    "madalena": "PE", "boa viagem": "PE", "imbiribeira": "PE",
    "peixinhos": "PE", "cabo de santo agostinho": "PE", "caxanga": "PE",
    "caxangá": "PE", "cordeiro": "PE", "encruzilhada": "PE",
    "hipódromo": "PE", "hipodromo": "PE", "torre": "PE",
    "apipucos": "PE", "dois irmaos": "PE", "dois irmãos": "PE",
    "vasco da gama": "PE",
    # ── Cidades/regiões PB ────────────────────────────────────────────────────
    "campina grande": "PB", "joão pessoa": "PB", "joao pessoa": "PB",
    "patos": "PB", "santa rita": "PB", "cabedelo": "PB",
    # ── Cidades/regiões RN ────────────────────────────────────────────────────
    "caicó": "RN", "caico": "RN", "mossoró": "RN", "mossoro": "RN",
    "natal": "RN", "parnamirim": "RN",
    # Bairros de Natal — RN
    "capim macio": "RN", "capim mácio": "RN", "candelaria": "RN",
    "candelária": "RN", "ponta negra": "RN", "lagoa nova": "RN",
    "tirol": "RN", "neópolis": "RN", "neopolis": "RN",
    "sao goncalo": "RN", "são gonçalo": "RN",
    # ── Cidades/regiões AL ────────────────────────────────────────────────────
    "arapiraca": "AL", "maceió": "AL", "maceio": "AL",
    # Bairros de Maceió — AL
    "mangabeiras": "AL", "gruta de lourdes": "AL", "jardim sao paulo": "AL",
    "jardim são paulo": "AL", "pajuçara": "AL", "pajucara": "AL",
    "jatiuca": "AL", "jatiúca": "AL", "ponta verde": "AL",
    "farol": "AL",
    # ── Cidades/regiões BA ────────────────────────────────────────────────────
    "alagoinhas": "BA", "barreiras": "BA",
    "cabula": "BA", "camaçari": "BA", "camacari": "BA",
    "cosme de farias": "BA", "eunápolis": "BA", "eunapolis": "BA",
    "feira de santana": "BA", "feira morada": "BA", "ilhéus": "BA",
    "ilheus": "BA", "irecê": "BA", "irece": "BA", "itapuã": "BA",
    "itapua": "BA", "juazeiro": "BA", "lauro de freitas": "BA",
    "mares": "BA", "pau da lima": "BA", "pituba": "BA", "salvador": "BA",
    "santo antônio de jesus": "BA", "santo antonio de jesus": "BA",
    "simões filho": "BA", "simoes filho": "BA", "teixeira de freitas": "BA",
    "valença": "BA", "valenca": "BA", "vitória da conquista": "BA",
    "vitoria da conquista": "BA",
    # Cidades do interior — BA
    "barris": "BA", "guanambi": "BA", "itabaiona": "BA", "itapetinga": "BA",
    "jequie": "BA", "jequié": "BA", "paulo afonso": "BA",
    "senhor do bonfim": "BA", "serrinha": "BA",
    # Bairros de Salvador — BA
    "barra": "BA", "garibaldi": "BA", "canela": "BA", "iguatemi": "BA",
    "acm": "BA", "brotas": "BA", "boca do rio": "BA", "imbuí": "BA",
    "imbui": "BA", "ondina": "BA", "liberdade": "BA", "engenho velho": "BA",
    "paralela": "BA", "aeroclube": "BA", "pernambués": "BA", "pernambuces": "BA",
    "federacao": "BA", "federação": "BA", "uruguai": "BA",
    "cidade baixa": "BA", "nazare": "BA", "nazaré": "BA",
    "ribeira": "BA", "barbalho": "BA", "largo do tanque": "BA",
    "mussurunga": "BA", "paripe": "BA", "jardim cajazeiras": "BA",
    "cajazeiras": "BA", "periperi": "BA", "periperi": "BA",
    "pirajá": "BA", "piraja": "BA", "são caetano": "BA", "sao caetano": "BA",
    # ── SE ───────────────────────────────────────────────────────────────────
    "aracaju": "SE", "nossa senhora do socorro": "SE", "socorro": "SE",
}


# Pré-normaliza as chaves do mapa para comparação sem acento
_MAPA_NORM = {_sem_acento(k): v for k, v in MAPA_ESTADO.items()}


def _estado_por_nome(nome):
    n = _sem_acento(nome)
    for kw, est in _MAPA_NORM.items():
        if kw in n:
            return est
    return ""


# ── NORMALIZAÇÃO ──────────────────────────────────────────────────────────────
def _norm_cod(v):
    if v is None:
        return ""
    try:
        return str(int(float(str(v).strip().lstrip("0") or "0")))
    except Exception:
        return str(v).strip()


def _norm_val(v):
    try:
        return float(str(v).replace(".", "").replace(",", "."))
    except Exception:
        return 0.0


def _norm_forn(v, fabricas):
    vl = _sem_acento(v).strip()
    for fab in fabricas:
        if _sem_acento(fab["nome_faturamento"]) in vl:
            return fab
    return None


# ── PARSE HTML-XLS ────────────────────────────────────────────────────────────
def _parse_html_xls(conteudo_bytes):
    """Lê arquivo .xls disfarçado de HTML e retorna lista de linhas."""
    try:
        html = conteudo_bytes.decode("latin-1", errors="replace")
    except Exception:
        html = conteudo_bytes.decode("utf-8", errors="replace")
    soup = BeautifulSoup(html, "lxml")
    rows = soup.find_all("tr")
    dados = []
    for r in rows[4:]:
        cells = [td.get_text(strip=True) for td in r.find_all(["td", "th"])]
        if len(cells) >= 10:
            dados.append(cells)
    return dados


# ── PARSE XLSX REAL ───────────────────────────────────────────────────────────
def _parse_xlsx(conteudo_bytes):
    df = pd.read_excel(io.BytesIO(conteudo_bytes), header=None)
    # Encontrar linha de cabeçalho (que contenha "Cliente" ou "Fornecedor")
    header_row = 0
    for i, row in df.iterrows():
        vals = [str(v).lower() for v in row.values]
        if any("cliente" in v for v in vals) and any("fornecedor" in v for v in vals):
            header_row = i
            break
    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    dados = []
    for _, row in df.iterrows():
        try:
            cells = [str(row.get(c, "")).strip() for c in df.columns]
            if len(cells) >= 10:
                dados.append(cells)
        except Exception:
            continue
    return dados


# ── DETECTAR REDE ─────────────────────────────────────────────────────────────
def _detectar_rede(nome_cliente, redes):
    """Retorna a rede correspondente ao nome do cliente, considerando estados."""
    nc = _sem_acento(nome_cliente)
    estado = _estado_por_nome(nome_cliente)
    for rede in redes:
        filtros = [_sem_acento(f.strip()) for f in rede["filtro_nome"].split(",") if f.strip()]
        if not any(f in nc for f in filtros):
            continue
        # Verificar exclusões
        excluir = [_sem_acento(e.strip()) for e in (rede["excluir_palavras"] or "").split(",") if e.strip()]
        if any(ex in nc for ex in excluir):
            continue
        # Verificar estados
        estados_rede = [e.strip() for e in (rede["estados"] or "").split(",") if e.strip()]
        if estados_rede:
            # Se o estado da loja não foi detectado E a rede exige estado → não casar
            # (evita que lojas com bairro desconhecido caiam em todas as redes)
            if not estado or estado not in estados_rede:
                continue
        return rede, estado
    return None, estado


# ── IMPORTAR FATURAMENTO ──────────────────────────────────────────────────────
def importar_faturamento(conteudo_bytes, nome_arquivo, callback_progresso=None):
    """
    Lê o faturamento, identifica fábrica/rede/produto e grava no banco.
    nome_arquivo no formato YYYY-MM (ex: 2026-06) define a data do período.
    Retorna dict com estatísticas.
    """
    # Data do período baseada no nome do arquivo (YYYY-MM → YYYY-MM-01)
    try:
        partes = nome_arquivo.split("-")
        data_periodo = f"{partes[0]}-{partes[1]}-01"
    except Exception:
        data_periodo = ""

    fabricas = banco.listar_fabricas()
    redes    = banco.listar_redes()

    # Descobrir produtos de todas as fábricas indexados por (fabrica_id, codigo_fab)
    todos_prods = banco.listar_produtos()
    idx_prods = {}
    for p in todos_prods:
        idx_prods[(p["fabrica_id"], p["codigo_fab"])] = p["id"]

    # Parse do arquivo
    ext = os.path.splitext(nome_arquivo)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        try:
            linhas = _parse_xlsx(conteudo_bytes)
        except Exception:
            linhas = _parse_html_xls(conteudo_bytes)
    else:
        try:
            linhas = _parse_html_xls(conteudo_bytes)
        except Exception:
            linhas = _parse_xlsx(conteudo_bytes)

    stats = {
        "total": len(linhas),
        "gravados": 0,
        "sem_fabrica": 0,
        "sem_rede": 0,
        "sem_produto": 0,
        "erros": 0,
        "erros_detalhe": [],   # lista de {linha, celulas, motivo}
    }

    # Limpar registros anteriores do mesmo arquivo
    banco.limpar_faturamento_arquivo(nome_arquivo)

    registros = []
    cache_lojas = {}  # nome_cliente -> loja_id
    datas = []

    for idx, cells in enumerate(linhas):
        try:
            if len(cells) < 10:
                continue
            # Colunas: Pedido, Cliente, Ped.Forn, Fornecedor, Vendedor, Status,
            #          Código, Produto, Valor, Qtd, Faturado, Total, Dupli, Total Geral
            nome_cliente = cells[1].strip()
            forn_raw     = cells[3].strip()
            cod_raw      = cells[6].strip()
            data_raw     = data_periodo  # data do período selecionado na importação
            qtd_raw      = cells[9].strip()
            valor_raw    = cells[11].strip()

            # Fábrica
            fab = _norm_forn(forn_raw, fabricas)
            if not fab:
                stats["sem_fabrica"] += 1
                continue

            # Rede
            rede, estado = _detectar_rede(nome_cliente, redes)
            if not rede:
                stats["sem_rede"] += 1
                continue

            # Produto
            cod_norm = _norm_cod(cod_raw)
            prod_id = idx_prods.get((fab["id"], cod_norm))
            if not prod_id:
                stats["sem_produto"] += 1
                continue

            # Loja
            if nome_cliente not in cache_lojas:
                cache_lojas[nome_cliente] = banco.obter_ou_criar_loja(
                    rede["id"], nome_cliente, estado
                )
            loja_id = cache_lojas[nome_cliente]

            qtd   = _norm_val(qtd_raw)
            valor = _norm_val(valor_raw)

            registros.append({
                "produto_id":     prod_id,
                "loja_id":        loja_id,
                "data_pedido":    data_raw,
                "qtd_vendida":    qtd,
                "valor_total":    valor,
                "arquivo_origem": nome_arquivo,
            })
            stats["gravados"] += 1

            if callback_progresso and idx % 500 == 0:
                callback_progresso(idx, len(linhas))

        except Exception as _exc:
            import traceback as _tb
            stats["erros"] += 1
            stats["erros_detalhe"].append({
                "linha":   idx + 1,
                "cliente": cells[1].strip() if len(cells) > 1 else "?",
                "motivo":  str(_exc) or type(_exc).__name__,
            })
            continue

    # Gravar em lote
    if registros:
        banco.gravar_faturamento(registros)

    # Registrar importação
    banco.registrar_importacao(
        arquivo=nome_arquivo,
        periodo_inicio="",
        periodo_fim="",
        n_linhas=stats["gravados"],
        status="ok" if stats["erros"] == 0 else f"{stats['erros']} erros",
    )

    return stats
