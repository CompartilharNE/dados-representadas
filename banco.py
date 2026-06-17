"""
banco.py — Operacoes do banco PostgreSQL (Supabase)
Compartilhar NE — Dados Representadas
v2
"""
import psycopg2
import psycopg2.extras
import streamlit as st

# Parametros fixos do Supabase Transaction Pooler
_DB_HOST = "aws-1-us-east-1.pooler.supabase.com"
_DB_PORT = 6543
_DB_NAME = "postgres"
_DB_USER = "postgres.qepdafhofbbuooinkieh"


def conectar():
    conn = psycopg2.connect(
        host=_DB_HOST,
        port=_DB_PORT,
        database=_DB_NAME,
        user=_DB_USER,
        password=st.secrets["database"]["password"],
        sslmode="require",
        connect_timeout=10,
    )
    return conn


def _fetch(sql, params=None, one=False):
    conn = conectar()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(sql, params or ())
            return cur.fetchone() if one else cur.fetchall()
    finally:
        conn.close()


def _run(sql, params=None, returning=False):
    conn = conectar()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params or ())
            result = cur.fetchone()[0] if returning else None
        conn.commit()
        return result
    finally:
        conn.close()


def criar_banco():
    conn = conectar()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS fabricas (
                    id SERIAL PRIMARY KEY,
                    nome TEXT NOT NULL UNIQUE,
                    nome_faturamento TEXT NOT NULL,
                    contato TEXT DEFAULT ''
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS redes (
                    id SERIAL PRIMARY KEY,
                    nome TEXT NOT NULL UNIQUE,
                    filtro_nome TEXT NOT NULL,
                    estados TEXT NOT NULL DEFAULT '',
                    excluir_palavras TEXT DEFAULT ''
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS produtos (
                    id SERIAL PRIMARY KEY,
                    fabrica_id INTEGER NOT NULL REFERENCES fabricas(id),
                    codigo_fab TEXT NOT NULL,
                    nome TEXT NOT NULL,
                    familia TEXT DEFAULT '',
                    peso TEXT DEFAULT '',
                    qtde_cx TEXT DEFAULT '',
                    und TEXT DEFAULT '',
                    preco TEXT DEFAULT '',
                    UNIQUE(fabrica_id, codigo_fab)
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS codigos_rede (
                    id SERIAL PRIMARY KEY,
                    produto_id INTEGER NOT NULL REFERENCES produtos(id),
                    rede_id INTEGER NOT NULL REFERENCES redes(id),
                    codigo_rede TEXT NOT NULL,
                    ativo INTEGER DEFAULT 1,
                    UNIQUE(produto_id, rede_id)
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS lojas (
                    id SERIAL PRIMARY KEY,
                    rede_id INTEGER REFERENCES redes(id),
                    nome_faturamento TEXT NOT NULL UNIQUE,
                    cidade TEXT DEFAULT '',
                    estado TEXT DEFAULT ''
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS faturamento (
                    id SERIAL PRIMARY KEY,
                    produto_id INTEGER REFERENCES produtos(id),
                    loja_id INTEGER REFERENCES lojas(id),
                    data_pedido TEXT DEFAULT '',
                    qtd_vendida DOUBLE PRECISION DEFAULT 0,
                    valor_total DOUBLE PRECISION DEFAULT 0,
                    arquivo_origem TEXT DEFAULT ''
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS importacoes (
                    id SERIAL PRIMARY KEY,
                    arquivo TEXT,
                    data_importacao TIMESTAMP DEFAULT NOW(),
                    periodo_inicio TEXT DEFAULT '',
                    periodo_fim TEXT DEFAULT '',
                    n_linhas INTEGER DEFAULT 0,
                    status TEXT DEFAULT 'ok'
                )
            """)

            # Só popula dados iniciais se as tabelas estiverem vazias (primeiro uso)
            cur.execute("SELECT COUNT(*) FROM fabricas")
            if cur.fetchone()[0] == 0:
                fabricas = [
                    ("Quata", "Nova Mix"),
                    ("Allfood/Granarolo", "Allfood"),
                    ("Bufalissima", "Bufalissima"),
                    ("Prieto", "Prieto"),
                    ("Sao Vicente", "Sao Vicente"),
                    ("Villa Germania", "Villa Germania"),
                ]
                for nome, nome_fat in fabricas:
                    cur.execute(
                        "INSERT INTO fabricas (nome, nome_faturamento) VALUES (%s, %s) ON CONFLICT (nome) DO NOTHING",
                        (nome, nome_fat)
                    )

            cur.execute("SELECT COUNT(*) FROM redes")
            if cur.fetchone()[0] == 0:
                redes_iniciais = [
                    ("Atacadao PE", "Atacadao", "PE,PB,RN,AL", "atacado da carne,atacadao canela,garibaldi"),
                    ("Atacadao BA", "Atacadao", "BA,SE",        "atacado da carne,atacadao canela,garibaldi"),
                    ("Assai",       "Assai",    "",              ""),
                    ("GBarbosa",    "GBarbosa", "",              ""),
                    ("Carrefour",   "Carrefour", "",             ""),
                ]
                for nome, filtro, estados, excluir in redes_iniciais:
                    cur.execute(
                        "INSERT INTO redes (nome, filtro_nome, estados, excluir_palavras) VALUES (%s, %s, %s, %s) ON CONFLICT (nome) DO NOTHING",
                        (nome, filtro, estados, excluir)
                    )
        conn.commit()
    finally:
        conn.close()


# ── FABRICAS ──────────────────────────────────────────────────────────────────

def listar_fabricas():
    rows = _fetch("SELECT * FROM fabricas ORDER BY nome")
    return [dict(r) for r in rows]


def salvar_fabrica(nome, nome_faturamento, contato="", fab_id=None):
    if fab_id:
        _run("UPDATE fabricas SET nome=%s, nome_faturamento=%s, contato=%s WHERE id=%s",
             (nome, nome_faturamento, contato, fab_id))
    else:
        _run("INSERT INTO fabricas (nome, nome_faturamento, contato) VALUES (%s, %s, %s)",
             (nome, nome_faturamento, contato))


def deletar_fabrica(fab_id):
    conn = conectar()
    conn.autocommit = True
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM faturamento WHERE produto_id IN (SELECT id FROM produtos WHERE fabrica_id=%s)", (fab_id,))
            cur.execute("DELETE FROM codigos_rede WHERE produto_id IN (SELECT id FROM produtos WHERE fabrica_id=%s)", (fab_id,))
            cur.execute("DELETE FROM produtos WHERE fabrica_id=%s", (fab_id,))
            cur.execute("DELETE FROM fabricas WHERE id=%s", (fab_id,))
    finally:
        conn.close()


# ── REDES ─────────────────────────────────────────────────────────────────────

def listar_redes():
    rows = _fetch("SELECT * FROM redes ORDER BY nome")
    return [dict(r) for r in rows]


def salvar_rede(nome, filtro_nome, estados, excluir_palavras="", rede_id=None):
    if rede_id:
        _run("UPDATE redes SET nome=%s, filtro_nome=%s, estados=%s, excluir_palavras=%s WHERE id=%s",
             (nome, filtro_nome, estados, excluir_palavras, rede_id))
    else:
        _run("INSERT INTO redes (nome, filtro_nome, estados, excluir_palavras) VALUES (%s, %s, %s, %s)",
             (nome, filtro_nome, estados, excluir_palavras))


def deletar_rede(rede_id):
    conn = conectar()
    conn.autocommit = True
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM faturamento WHERE loja_id IN (SELECT id FROM lojas WHERE rede_id=%s)", (rede_id,))
            cur.execute("DELETE FROM codigos_rede WHERE rede_id=%s", (rede_id,))
            cur.execute("DELETE FROM lojas WHERE rede_id=%s", (rede_id,))
            cur.execute("DELETE FROM redes WHERE id=%s", (rede_id,))
    finally:
        conn.close()


# ── PRODUTOS ──────────────────────────────────────────────────────────────────

def listar_produtos(fabrica_id=None):
    if fabrica_id:
        rows = _fetch(
            "SELECT p.*, f.nome as fabrica_nome FROM produtos p "
            "JOIN fabricas f ON f.id=p.fabrica_id WHERE p.fabrica_id=%s ORDER BY p.familia, p.nome",
            (fabrica_id,)
        )
    else:
        rows = _fetch(
            "SELECT p.*, f.nome as fabrica_nome FROM produtos p "
            "JOIN fabricas f ON f.id=p.fabrica_id ORDER BY f.nome, p.familia, p.nome"
        )
    return [dict(r) for r in rows]


def contar_produtos(fabrica_id):
    row = _fetch("SELECT COUNT(*) as cnt FROM produtos WHERE fabrica_id=%s", (fabrica_id,), one=True)
    return row["cnt"] if row else 0


def salvar_produto(fabrica_id, codigo_fab, nome, familia="", peso="", qtde_cx="", und="", preco=""):
    _run("""
        INSERT INTO produtos (fabrica_id, codigo_fab, nome, familia, peso, qtde_cx, und, preco)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (fabrica_id, codigo_fab) DO UPDATE SET
            nome=EXCLUDED.nome, familia=EXCLUDED.familia, peso=EXCLUDED.peso,
            qtde_cx=EXCLUDED.qtde_cx, und=EXCLUDED.und, preco=EXCLUDED.preco
    """, (fabrica_id, str(codigo_fab), nome, familia, peso, str(qtde_cx), und, preco))


def importar_catalogo_df(fabrica_id, df, col_map=None):
    """
    col_map: dict com chaves (codigo_fab, nome, familia, peso, qtde_cx, und, preco)
             mapeando para nomes de colunas do df. Valores None = ignorar campo.
    Se col_map não fornecido, tenta detectar por nome de coluna (legado).
    """
    def _get(row, campo, fallbacks):
        if col_map and col_map.get(campo):
            val = row.get(col_map[campo], "")
        else:
            val = ""
            for fb in fallbacks:
                val = row.get(fb, "")
                if val and str(val) not in ("nan", "None", ""):
                    break
        v = str(val).strip()
        return "" if v in ("nan", "None") else v

    conn = conectar()
    n = 0
    try:
        with conn.cursor() as cur:
            for _, row in df.iterrows():
                cod  = _get(row, "codigo_fab", ["Codigo", "COD", "Cód", "código"])
                nome = _get(row, "nome",       ["Produto", "PRODUTO", "Nome", "NOME", "Descrição"])
                if not cod or not nome:
                    continue
                try:
                    cur.execute("""
                        INSERT INTO produtos (fabrica_id, codigo_fab, nome, familia, peso, qtde_cx, und, preco)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                        ON CONFLICT (fabrica_id, codigo_fab) DO UPDATE SET
                            nome=EXCLUDED.nome, familia=EXCLUDED.familia, peso=EXCLUDED.peso,
                            qtde_cx=EXCLUDED.qtde_cx, und=EXCLUDED.und, preco=EXCLUDED.preco
                    """, (
                        fabrica_id,
                        cod,
                        nome,
                        _get(row, "familia",  ["Familia", "FAMILIA", "Categoria", "Linha"]),
                        _get(row, "peso",     ["Peso", "PESO", "Gramatura"]),
                        _get(row, "qtde_cx",  ["Qtd", "QTD", "QtdCX", "CX"]),
                        _get(row, "und",      ["UND", "Unidade", "Embalagem"]),
                        _get(row, "preco",    ["Preco", "PRECO", "Preço", "Valor"]),
                    ))
                    n += 1
                except Exception:
                    continue
        conn.commit()
    finally:
        conn.close()
    return n


# ── CODIGOS DE REDE ───────────────────────────────────────────────────────────

def listar_codigos_rede(fabrica_id, rede_id):
    rows = _fetch("""
        SELECT cr.*, p.codigo_fab, p.nome as produto_nome
        FROM codigos_rede cr
        JOIN produtos p ON p.id = cr.produto_id
        WHERE p.fabrica_id=%s AND cr.rede_id=%s
        ORDER BY p.familia, p.nome
    """, (fabrica_id, rede_id))
    return [dict(r) for r in rows]


def salvar_codigo_rede(produto_id, rede_id, codigo_rede, ativo=1):
    _run("""
        INSERT INTO codigos_rede (produto_id, rede_id, codigo_rede, ativo)
        VALUES (%s, %s, %s, %s)
        ON CONFLICT (produto_id, rede_id) DO UPDATE SET
            codigo_rede=EXCLUDED.codigo_rede, ativo=EXCLUDED.ativo
    """, (produto_id, rede_id, str(codigo_rede), ativo))


def importar_codigos_rede_df(fabrica_id, rede_id, df):
    conn = conectar()
    n = 0
    nao_encontrados = []
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            for _, row in df.iterrows():
                cod_fab  = str(row.get("codigo_fab",  row.get("COD FAB",    row.get("QUATA", "")))).strip()
                cod_rede = str(row.get("codigo_rede", row.get("COD ATACADAO", row.get("ATACADAO", "")))).strip()
                if not cod_fab or not cod_rede or cod_fab in ("nan", "None") or cod_rede in ("nan", "None"):
                    continue
                try:
                    cod_fab_norm = str(int(float(cod_fab)))
                except Exception:
                    cod_fab_norm = cod_fab
                cur.execute(
                    "SELECT id FROM produtos WHERE fabrica_id=%s AND codigo_fab=%s",
                    (fabrica_id, cod_fab_norm)
                )
                prod = cur.fetchone()
                if not prod:
                    nao_encontrados.append(cod_fab_norm)
                    continue
                cur.execute("""
                    INSERT INTO codigos_rede (produto_id, rede_id, codigo_rede)
                    VALUES (%s, %s, %s)
                    ON CONFLICT (produto_id, rede_id) DO UPDATE SET codigo_rede=EXCLUDED.codigo_rede
                """, (prod["id"], rede_id, cod_rede))
                n += 1
        conn.commit()
    finally:
        conn.close()
    return n, nao_encontrados


# ── LOJAS ─────────────────────────────────────────────────────────────────────

def listar_lojas(rede_id=None):
    if rede_id:
        rows = _fetch(
            "SELECT * FROM lojas WHERE rede_id=%s ORDER BY estado, nome_faturamento",
            (rede_id,)
        )
    else:
        rows = _fetch("SELECT * FROM lojas ORDER BY estado, nome_faturamento")
    return [dict(r) for r in rows]


def obter_ou_criar_loja(rede_id, nome_faturamento, estado=""):
    conn = conectar()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM lojas WHERE nome_faturamento=%s", (nome_faturamento,))
            loja = cur.fetchone()
            if loja:
                return loja[0]
            cur.execute(
                "INSERT INTO lojas (rede_id, nome_faturamento, estado) VALUES (%s, %s, %s) RETURNING id",
                (rede_id, nome_faturamento, estado)
            )
            loja_id = cur.fetchone()[0]
        conn.commit()
        return loja_id
    finally:
        conn.close()


# ── FATURAMENTO ───────────────────────────────────────────────────────────────

def gravar_faturamento(registros):
    if not registros:
        return
    conn = conectar()
    try:
        with conn.cursor() as cur:
            psycopg2.extras.execute_values(cur, """
                INSERT INTO faturamento (produto_id, loja_id, data_pedido, qtd_vendida, valor_total, arquivo_origem)
                VALUES %s
            """, [(
                r["produto_id"], r["loja_id"], r["data_pedido"],
                r["qtd_vendida"], r["valor_total"], r["arquivo_origem"]
            ) for r in registros])
        conn.commit()
    finally:
        conn.close()


def limpar_faturamento_arquivo(arquivo):
    _run("DELETE FROM faturamento WHERE arquivo_origem=%s", (arquivo,))


def limpar_todo_faturamento():
    conn = conectar()
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM faturamento")
            cur.execute("DELETE FROM importacoes")
        conn.commit()
    finally:
        conn.close()


def listar_importacoes():
    rows = _fetch(
        "SELECT id, arquivo, TO_CHAR(data_importacao, 'YYYY-MM-DD HH24:MI:SS') as data_importacao, "
        "periodo_inicio, periodo_fim, n_linhas, status FROM importacoes ORDER BY data_importacao DESC LIMIT 20"
    )
    return [dict(r) for r in rows]


def registrar_importacao(arquivo, periodo_inicio, periodo_fim, n_linhas, status="ok"):
    _run("""
        INSERT INTO importacoes (arquivo, periodo_inicio, periodo_fim, n_linhas, status)
        VALUES (%s, %s, %s, %s, %s)
    """, (arquivo, periodo_inicio, periodo_fim, n_linhas, status))


# ── CONSULTAS PARA RELATORIO ──────────────────────────────────────────────────

def obter_fabrica_por_nome(nome):
    row = _fetch("SELECT * FROM fabricas WHERE nome=%s", (nome,), one=True)
    return dict(row) if row else None


def obter_rede_por_nome(nome):
    row = _fetch("SELECT * FROM redes WHERE nome=%s", (nome,), one=True)
    return dict(row) if row else None


def produtos_com_codigos(fabrica_id, rede_id):
    rows = _fetch("""
        SELECT p.*, cr.codigo_rede
        FROM produtos p
        JOIN codigos_rede cr ON cr.produto_id=p.id
        WHERE p.fabrica_id=%s AND cr.rede_id=%s AND cr.ativo=1
        ORDER BY p.familia, p.nome
    """, (fabrica_id, rede_id))
    return [dict(r) for r in rows]


def vendas_por_produto(fabrica_id, rede_id, data_inicio, data_fim):
    rows = _fetch("""
        SELECT f.produto_id,
               SUM(f.qtd_vendida) as qtd_total,
               SUM(f.valor_total) as valor_total,
               COUNT(DISTINCT f.loja_id) as n_lojas
        FROM faturamento f
        JOIN lojas l ON l.id = f.loja_id
        WHERE l.rede_id=%s
          AND f.produto_id IN (SELECT p.id FROM produtos p WHERE p.fabrica_id=%s)
          AND (%s = '' OR f.data_pedido >= %s)
          AND (%s = '' OR f.data_pedido <= %s)
        GROUP BY f.produto_id
    """, (rede_id, fabrica_id, data_inicio, data_inicio, data_fim, data_fim))
    return {r["produto_id"]: dict(r) for r in rows}


def vendas_por_loja(fabrica_id, rede_id, data_inicio, data_fim):
    rows = _fetch("""
        SELECT l.nome_faturamento, l.estado,
               SUM(f.valor_total) as valor_total,
               COUNT(DISTINCT f.produto_id) as n_produtos
        FROM faturamento f
        JOIN lojas l ON l.id = f.loja_id
        WHERE l.rede_id=%s
          AND f.produto_id IN (SELECT p.id FROM produtos p WHERE p.fabrica_id=%s)
          AND (%s = '' OR f.data_pedido >= %s)
          AND (%s = '' OR f.data_pedido <= %s)
        GROUP BY f.loja_id, l.nome_faturamento, l.estado
        ORDER BY valor_total DESC
    """, (rede_id, fabrica_id, data_inicio, data_inicio, data_fim, data_fim))
    return [dict(r) for r in rows]
