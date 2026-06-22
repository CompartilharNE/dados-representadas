"""
banco.py — Operacoes do banco PostgreSQL (Supabase)
Compartilhar NE — Dados Representadas
v2.1
"""
import psycopg2
import psycopg2.extras
import streamlit as st
import bcrypt

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
                CREATE TABLE IF NOT EXISTS usuarios (
                    id SERIAL PRIMARY KEY,
                    username TEXT NOT NULL UNIQUE,
                    nome TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    perfil TEXT NOT NULL DEFAULT 'usuario',
                    ultimo_acesso TIMESTAMP
                )
            """)
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
            # Limpar duplicatas em lojas antes de criar o UNIQUE index
            cur.execute("""
                DELETE FROM lojas
                WHERE id IN (
                    SELECT id FROM (
                        SELECT id,
                               ROW_NUMBER() OVER (
                                   PARTITION BY LOWER(nome_faturamento)
                                   ORDER BY id
                               ) AS rn
                        FROM lojas
                    ) sub
                    WHERE rn > 1
                )
                  AND id NOT IN (SELECT DISTINCT loja_id FROM faturamento WHERE loja_id IS NOT NULL)
            """)
            # Reatribuir faturamento de duplicatas antes de removê-las
            cur.execute("""
                UPDATE faturamento f
                SET loja_id = (
                    SELECT MIN(l2.id)
                    FROM lojas l2
                    WHERE LOWER(l2.nome_faturamento) = LOWER(
                        (SELECT nome_faturamento FROM lojas WHERE id = f.loja_id)
                    )
                )
                WHERE f.loja_id IN (
                    SELECT id FROM lojas
                    WHERE LOWER(nome_faturamento) IN (
                        SELECT LOWER(nome_faturamento)
                        FROM lojas
                        GROUP BY LOWER(nome_faturamento)
                        HAVING COUNT(*) > 1
                    )
                    AND id NOT IN (
                        SELECT MIN(id)
                        FROM lojas
                        GROUP BY LOWER(nome_faturamento)
                    )
                )
            """)
            cur.execute("""
                DELETE FROM lojas
                WHERE id NOT IN (
                    SELECT MIN(id)
                    FROM lojas
                    GROUP BY LOWER(nome_faturamento)
                )
            """)
            # Garantir UNIQUE index em lojas.nome_faturamento (necessário para ON CONFLICT)
            cur.execute("""
                CREATE UNIQUE INDEX IF NOT EXISTS idx_lojas_nome_fat
                ON lojas (nome_faturamento)
            """)
            # Adicionar coluna logo_b64 à tabela redes (se ainda não existir)
            cur.execute("""
                ALTER TABLE redes ADD COLUMN IF NOT EXISTS logo_b64 TEXT DEFAULT ''
            """)
            # Adicionar coluna logo_b64 à tabela fabricas (se ainda não existir)
            cur.execute("""
                ALTER TABLE fabricas ADD COLUMN IF NOT EXISTS logo_b64 TEXT DEFAULT ''
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS codigos_fornecedor (
                    id SERIAL PRIMARY KEY,
                    fabrica_id INTEGER NOT NULL REFERENCES fabricas(id),
                    rede_id    INTEGER NOT NULL REFERENCES redes(id),
                    codigo     TEXT NOT NULL DEFAULT '',
                    UNIQUE(fabrica_id, rede_id)
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS tabela_precos (
                    id SERIAL PRIMARY KEY,
                    rede_id INTEGER NOT NULL REFERENCES redes(id),
                    fabrica_id INTEGER NOT NULL REFERENCES fabricas(id),
                    produto_id INTEGER REFERENCES produtos(id),
                    codigo_fab TEXT NOT NULL,
                    descricao TEXT DEFAULT '',
                    unidade TEXT DEFAULT '',
                    preco NUMERIC(12,4) DEFAULT 0,
                    data_atualizacao TIMESTAMP DEFAULT NOW(),
                    UNIQUE(rede_id, fabrica_id, codigo_fab)
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

    # Corrige estados vazios de lojas já importadas usando o mapa atual
    try:
        corrigir_estados_lojas()
    except Exception:
        pass  # seguro: só atualiza registros existentes

    # Correções pontuais de lojas mal-classificadas
    try:
        corrigir_loja_rede_por_nome(
            "Sendas Federação (Av. Vasco da Gama) (256)", "Assai BA", "BA"
        )
    except Exception:
        pass


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


def salvar_logo_rede(rede_id, logo_b64):
    """Salva a logo da rede (base64) no banco."""
    _run("UPDATE redes SET logo_b64=%s WHERE id=%s", (logo_b64, rede_id))


def salvar_logo_fabrica(fab_id, logo_b64):
    """Salva a logo da fábrica (base64) no banco."""
    _run("UPDATE fabricas SET logo_b64=%s WHERE id=%s", (logo_b64, fab_id))


def salvar_cod_fornecedor(fabrica_id, rede_id, codigo):
    """Salva ou atualiza o código de fornecedor de uma fábrica numa rede."""
    _run("""
        INSERT INTO codigos_fornecedor (fabrica_id, rede_id, codigo)
        VALUES (%s, %s, %s)
        ON CONFLICT (fabrica_id, rede_id)
        DO UPDATE SET codigo = EXCLUDED.codigo
    """, (fabrica_id, rede_id, codigo))


def get_cod_fornecedor(fabrica_id, rede_id):
    """Retorna o código de fornecedor de uma fábrica numa rede, ou ''."""
    row = _fetch(
        "SELECT codigo FROM codigos_fornecedor WHERE fabrica_id=%s AND rede_id=%s",
        (fabrica_id, rede_id), one=True
    )
    return row["codigo"] if row else ""


def listar_cods_forn_por_rede(rede_id):
    """Retorna todos os códigos de fornecedor registrados para uma rede."""
    rows = _fetch("""
        SELECT cf.codigo, f.nome as fabrica_nome, cf.fabrica_id
        FROM codigos_fornecedor cf
        JOIN fabricas f ON f.id = cf.fabrica_id
        WHERE cf.rede_id = %s
        ORDER BY f.nome
    """, (rede_id,))
    return [dict(r) for r in rows]


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
    """Busca loja pelo nome (case-insensitive). Se existir com rede diferente, atualiza a rede.
    Usa ON CONFLICT para evitar UniqueViolation em inserções concorrentes.
    """
    conn = conectar()
    try:
        with conn.cursor() as cur:
            # Busca apenas pelo nome (UNIQUE é só em nome_faturamento)
            cur.execute(
                "SELECT id, rede_id FROM lojas WHERE LOWER(nome_faturamento)=LOWER(%s)",
                (nome_faturamento,)
            )
            loja = cur.fetchone()
            if loja:
                loja_id, rede_atual = loja
                # Atualiza rede e estado se a planilha indica rede diferente
                if rede_atual != rede_id or (estado and estado != ""):
                    cur.execute(
                        "UPDATE lojas SET rede_id=%s, estado=COALESCE(NULLIF(%s,''), estado) WHERE id=%s",
                        (rede_id, estado, loja_id)
                    )
                conn.commit()
                return loja_id
            # Insere; ON CONFLICT garante que concorrência não quebra
            cur.execute("""
                INSERT INTO lojas (rede_id, nome_faturamento, estado)
                VALUES (%s, %s, %s)
                ON CONFLICT (nome_faturamento) DO UPDATE
                    SET rede_id=EXCLUDED.rede_id,
                        estado=COALESCE(NULLIF(EXCLUDED.estado,''), lojas.estado)
                RETURNING id
            """, (rede_id, nome_faturamento, estado))
            loja_id = cur.fetchone()[0]
        conn.commit()
        return loja_id
    finally:
        conn.close()


def corrigir_estados_lojas():
    """Re-escaneia nomes de lojas com estado vazio e atualiza usando o mapa de estados."""
    import importar as _imp
    lojas = _fetch("SELECT id, nome_faturamento FROM lojas WHERE estado = '' OR estado IS NULL")
    atualizados = 0
    for l in lojas:
        est = _imp._estado_por_nome(l["nome_faturamento"])
        if est:
            _run("UPDATE lojas SET estado=%s WHERE id=%s", (est, l["id"]))
            atualizados += 1
    return atualizados


def corrigir_uf_por_rede():
    """
    Para lojas com UF vazia, tenta derivar pelo nome/estados da rede associada.
    Regra 1: últimas 2 letras maiúsculas do nome da rede (ex: 'Assai BA' → 'BA').
    Regra 2: se a rede tiver exatamente 1 estado configurado, usa esse estado.
    Retorna número de lojas atualizadas.
    """
    conn = conectar()
    atualizados = 0
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("""
                SELECT l.id, r.nome as rede_nome, r.estados
                FROM lojas l
                LEFT JOIN redes r ON r.id = l.rede_id
                WHERE (l.estado = '' OR l.estado IS NULL)
                  AND l.rede_id IS NOT NULL
            """)
            lojas = cur.fetchall()
            for loja in lojas:
                uf = ""
                rede_nome = loja["rede_nome"] or ""
                estados   = loja["estados"]   or ""
                # Regra 1: últimas 2 letras do nome da rede
                partes = rede_nome.strip().split()
                if partes and len(partes[-1]) == 2 and partes[-1].isupper():
                    uf = partes[-1]
                # Regra 2: rede com estado único
                if not uf:
                    lista = [e.strip() for e in estados.split(",") if e.strip()]
                    if len(lista) == 1:
                        uf = lista[0]
                if uf:
                    cur.execute("UPDATE lojas SET estado=%s WHERE id=%s", (uf, loja["id"]))
                    atualizados += 1
        conn.commit()
    finally:
        conn.close()
    return atualizados


def listar_lojas_com_rede(rede_id=None):
    """Retorna lojas com nome da rede, opcionalmente filtrado por rede_id."""
    if rede_id:
        rows = _fetch("""
            SELECT l.*, r.nome as rede_nome
            FROM lojas l
            LEFT JOIN redes r ON r.id = l.rede_id
            WHERE l.rede_id=%s
            ORDER BY l.estado, l.nome_faturamento
        """, (rede_id,))
    else:
        rows = _fetch("""
            SELECT l.*, r.nome as rede_nome
            FROM lojas l
            LEFT JOIN redes r ON r.id = l.rede_id
            ORDER BY r.nome, l.estado, l.nome_faturamento
        """)
    return [dict(r) for r in rows]


def _norm_str(s):
    """Normaliza string: minúsculas, sem acentos."""
    import unicodedata
    return unicodedata.normalize("NFD", str(s).lower()).encode("ascii", "ignore").decode("ascii").strip()


def importar_lojas_planilha(df, col_nome, col_uf, col_rede):
    """
    Importa lojas de um DataFrame com colunas Nome, UF e Rede.
    UPSERT: se a loja já existe, SUBSTITUI rede_id e estado.
    Retorna (n_ok, redes_nao_encontradas, erros, log_detalhado).
    log_detalhado = lista de dicts com info de cada linha processada.
    """
    redes = listar_redes()
    # Mapa sem acento para matching robusto (Assaí = Assai)
    redes_map = {_norm_str(r["nome"]): r for r in redes}

    def _match_rede(nome_planilha):
        n = _norm_str(nome_planilha)
        if n in redes_map:
            return redes_map[n]
        for k, v in redes_map.items():
            if n in k or k in n:
                return v
        return None

    n_ok = 0
    nao_enc = set()
    erros = []
    log = []  # diagnóstico linha a linha

    conn = conectar()
    try:
        with conn.cursor() as cur:
            for _, row in df.iterrows():
                nome  = str(row.get(col_nome, "") or "").strip()
                uf    = str(row.get(col_uf,   "") or "").strip().upper()
                r_nom = str(row.get(col_rede,  "") or "").strip()

                if not nome or nome in ("nan", "None", ""):
                    continue

                rede = _match_rede(r_nom)
                if not rede:
                    if r_nom and r_nom not in ("nan", "None"):
                        nao_enc.add(r_nom)
                    log.append({"nome": nome, "status": "rede_nao_encontrada", "rede_planilha": r_nom})
                    continue

                if uf in ("nan", "None"):
                    uf = ""

                # Se UF ainda vazia, tenta derivar do nome da rede (ex: "Assai BA" → "BA")
                if not uf:
                    partes = rede["nome"].strip().split()
                    if partes and len(partes[-1]) == 2 and partes[-1].isupper():
                        uf = partes[-1]
                # Se ainda vazia, usa estado da rede quando ela tem só um estado configurado
                if not uf:
                    estados_rede = [e.strip() for e in (rede.get("estados") or "").split(",") if e.strip()]
                    if len(estados_rede) == 1:
                        uf = estados_rede[0]

                try:
                    # Matching case-insensitive (evita duplicatas por diferença de maiúsculas)
                    cur.execute(
                        "SELECT id, nome_faturamento, estado FROM lojas WHERE LOWER(nome_faturamento)=LOWER(%s)",
                        (nome,)
                    )
                    existing = cur.fetchone()
                    if existing:
                        old_estado = existing[2] or ""
                        cur.execute(
                            "UPDATE lojas SET rede_id=%s, estado=%s, nome_faturamento=%s WHERE id=%s",
                            (rede["id"], uf if uf else old_estado, nome, existing[0])
                        )
                        log.append({"nome": nome, "status": "atualizado",
                                    "rede_planilha": r_nom, "rede_sistema": rede["nome"],
                                    "nome_banco": existing[1], "uf": uf})
                    else:
                        cur.execute(
                            "INSERT INTO lojas (rede_id, nome_faturamento, estado) VALUES (%s,%s,%s)",
                            (rede["id"], nome, uf)
                        )
                        log.append({"nome": nome, "status": "inserido",
                                    "rede_planilha": r_nom, "rede_sistema": rede["nome"], "uf": uf})
                    n_ok += 1
                except Exception as e:
                    erros.append(str(e))
                    log.append({"nome": nome, "status": "erro", "erro": str(e)})
        conn.commit()
    finally:
        conn.close()

    return n_ok, sorted(nao_enc), erros, log


def deletar_loja(loja_id):
    """Remove uma loja (e seus registros de faturamento vinculados)."""
    conn = conectar()
    conn.autocommit = True
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM faturamento WHERE loja_id=%s", (loja_id,))
            cur.execute("DELETE FROM lojas WHERE id=%s", (loja_id,))
    finally:
        conn.close()


def salvar_loja_manual(rede_id, nome_faturamento, uf):
    """Cria loja ou atualiza se já existir (por nome_faturamento)."""
    conn = conectar()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM lojas WHERE nome_faturamento=%s", (nome_faturamento,))
            ex = cur.fetchone()
            if ex:
                cur.execute(
                    "UPDATE lojas SET rede_id=%s, estado=%s WHERE id=%s",
                    (rede_id, uf, ex[0])
                )
            else:
                cur.execute(
                    "INSERT INTO lojas (rede_id, nome_faturamento, estado) VALUES (%s,%s,%s)",
                    (rede_id, nome_faturamento, uf)
                )
        conn.commit()
    finally:
        conn.close()


def atualizar_loja(loja_id, rede_id, nome_faturamento, uf):
    """Atualiza dados de uma loja existente.
    Se já existe outra loja com o mesmo nome, mescla (move faturamento e deleta a duplicata).
    """
    conn = conectar()
    try:
        with conn.cursor() as cur:
            # Verifica se outro registro já tem esse nome_faturamento
            cur.execute(
                "SELECT id FROM lojas WHERE LOWER(nome_faturamento)=LOWER(%s) AND id<>%s",
                (nome_faturamento, loja_id)
            )
            dup = cur.fetchone()
            if dup:
                dup_id = dup[0]
                # Move faturamento da duplicata para esta loja
                cur.execute(
                    "UPDATE faturamento SET loja_id=%s WHERE loja_id=%s",
                    (loja_id, dup_id)
                )
                # Remove a duplicata
                cur.execute("DELETE FROM lojas WHERE id=%s", (dup_id,))
            # Agora atualiza normalmente
            cur.execute(
                "UPDATE lojas SET rede_id=%s, nome_faturamento=%s, estado=%s WHERE id=%s",
                (rede_id, nome_faturamento, uf, loja_id)
            )
        conn.commit()
    finally:
        conn.close()


def limpar_lojas_duplicadas():
    """Remove lojas com nome_faturamento duplicado (LOWER), mantendo a de menor id.
    Move faturamento das duplicatas para a loja mantida antes de deletar.
    Retorna número de duplicatas removidas.
    """
    conn = conectar()
    removidos = 0
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            # Encontra grupos duplicados por nome (case-insensitive)
            cur.execute("""
                SELECT LOWER(nome_faturamento) AS nome_lower,
                       array_agg(id ORDER BY id) AS ids
                FROM lojas
                GROUP BY LOWER(nome_faturamento)
                HAVING count(*) > 1
            """)
            grupos = cur.fetchall()
            for g in grupos:
                ids = g["ids"]
                manter = ids[0]          # menor id = mais antigo
                remover = ids[1:]
                for dup_id in remover:
                    # Move faturamento
                    cur.execute(
                        "UPDATE faturamento SET loja_id=%s WHERE loja_id=%s",
                        (manter, dup_id)
                    )
                    cur.execute("DELETE FROM lojas WHERE id=%s", (dup_id,))
                    removidos += 1
        conn.commit()
    finally:
        conn.close()
    return removidos



def remover_prefixo_cliente():
    """Remove o prefixo 'Cliente: ' (case-insensitive) do nome_faturamento de todas as lojas.
    Retorna número de lojas atualizadas.
    """
    conn = conectar()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                UPDATE lojas
                SET nome_faturamento = TRIM(SUBSTRING(nome_faturamento FROM 9))
                WHERE nome_faturamento ILIKE 'Cliente: %'
            """)
            n = cur.rowcount
        conn.commit()
    finally:
        conn.close()
    return n



def corrigir_loja_rede_por_nome(nome_faturamento, nome_rede, estado):
    """Corrige rede e estado de uma loja específica pelo nome exato (case-insensitive).
    Usado para correções pontuais de lojas mal-classificadas.
    """
    conn = conectar()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM redes WHERE LOWER(nome)=LOWER(%s)", (nome_rede,))
            r = cur.fetchone()
            if not r:
                return False, f"Rede '{nome_rede}' não encontrada"
            rede_id = r[0]
            cur.execute(
                "UPDATE lojas SET rede_id=%s, estado=%s WHERE LOWER(nome_faturamento)=LOWER(%s)",
                (rede_id, estado, nome_faturamento)
            )
            n = cur.rowcount
        conn.commit()
        return True, n
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


def diagnostico_faturamento():
    """Retorna estatísticas do faturamento para diagnóstico."""
    rows = _fetch("""
        SELECT fab.nome as fabrica, r.nome as rede,
               COUNT(*) as registros,
               SUM(f.valor_total) as valor_total,
               MIN(f.data_pedido) as data_min,
               MAX(f.data_pedido) as data_max,
               f.arquivo_origem
        FROM faturamento f
        JOIN produtos p ON p.id = f.produto_id
        JOIN fabricas fab ON fab.id = p.fabrica_id
        JOIN lojas l ON l.id = f.loja_id
        JOIN redes r ON r.id = l.rede_id
        GROUP BY fab.nome, r.nome, f.arquivo_origem
        ORDER BY fab.nome, r.nome
    """)
    return [dict(r) for r in rows]


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


def todos_produtos_fabrica(fabrica_id):
    """Retorna todos os produtos da fábrica sem filtro de código de rede."""
    rows = _fetch("""
        SELECT p.*, '' as codigo_rede
        FROM produtos p
        WHERE p.fabrica_id=%s
        ORDER BY p.familia, p.nome
    """, (fabrica_id,))
    return [dict(r) for r in rows]


def ultimos_precos(fabrica_id, rede_id):
    """Retorna o último preço unitário faturado por produto_id para a rede."""
    rows = _fetch("""
        SELECT DISTINCT ON (f.produto_id)
            f.produto_id,
            ROUND((f.valor_total / NULLIF(f.qtd_vendida, 0))::numeric, 2) as preco_unitario
        FROM faturamento f
        JOIN lojas l ON l.id = f.loja_id
        WHERE l.rede_id = %s
          AND f.produto_id IN (SELECT id FROM produtos WHERE fabrica_id = %s)
          AND f.qtd_vendida > 0
        ORDER BY f.produto_id, f.data_pedido DESC, f.id DESC
    """, (rede_id, fabrica_id))
    return {r["produto_id"]: float(r["preco_unitario"]) for r in rows if r["preco_unitario"]}


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
          AND (f.data_pedido = '' OR %s = '' OR f.data_pedido >= %s)
          AND (f.data_pedido = '' OR %s = '' OR f.data_pedido <= %s)
        GROUP BY f.produto_id
    """, (rede_id, fabrica_id, data_inicio, data_inicio, data_fim, data_fim))
    return {r["produto_id"]: dict(r) for r in rows}


def vendas_por_loja(fabrica_id, rede_id, data_inicio, data_fim):
    """Retorna TODAS as lojas da rede, inclusive as que não compraram no período (valor_total=0)."""
    rows = _fetch("""
        SELECT l.id as loja_id, l.nome_faturamento, l.estado,
               COALESCE(SUM(f.valor_total), 0) as valor_total,
               COUNT(DISTINCT CASE WHEN f.valor_total > 0 THEN f.produto_id END) as n_produtos
        FROM lojas l
        LEFT JOIN faturamento f ON f.loja_id = l.id
          AND f.produto_id IN (
              SELECT cr.produto_id FROM codigos_rede cr
              WHERE cr.rede_id=%s AND cr.ativo=1
                AND cr.produto_id IN (SELECT id FROM produtos WHERE fabrica_id=%s)
          )
          AND (f.data_pedido = '' OR %s = '' OR f.data_pedido >= %s)
          AND (f.data_pedido = '' OR %s = '' OR f.data_pedido <= %s)
        WHERE l.rede_id=%s
        GROUP BY l.id, l.nome_faturamento, l.estado
        ORDER BY valor_total DESC, l.nome_faturamento
    """, (rede_id, fabrica_id, data_inicio, data_inicio, data_fim, data_fim, rede_id))
    return [dict(r) for r in rows]


# ── TABELA DE PREÇOS ─────────────────────────────────────────────────────────

def salvar_tabela_precos(rede_id, fabrica_id, items):
    """
    Salva/substitui tabela de preços para (rede_id, fabrica_id).
    items: lista de dicts com {codigo_fab, descricao, unidade, preco}
    """
    conn = conectar()
    try:
        with conn.cursor() as cur:
            # Remove entradas antigas desta rede+fábrica
            cur.execute(
                "DELETE FROM tabela_precos WHERE rede_id=%s AND fabrica_id=%s",
                (rede_id, fabrica_id)
            )
            for it in items:
                # Tenta encontrar o produto_id pelo codigo_fab
                cur.execute(
                    "SELECT id FROM produtos WHERE fabrica_id=%s AND codigo_fab=%s LIMIT 1",
                    (fabrica_id, str(it["codigo_fab"]).strip())
                )
                row = cur.fetchone()
                produto_id = row[0] if row else None
                cur.execute("""
                    INSERT INTO tabela_precos
                        (rede_id, fabrica_id, produto_id, codigo_fab, descricao, unidade, preco)
                    VALUES (%s,%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (rede_id, fabrica_id, codigo_fab)
                    DO UPDATE SET produto_id=EXCLUDED.produto_id,
                                  descricao=EXCLUDED.descricao,
                                  unidade=EXCLUDED.unidade,
                                  preco=EXCLUDED.preco,
                                  data_atualizacao=NOW()
                """, (
                    rede_id, fabrica_id, produto_id,
                    str(it["codigo_fab"]).strip(),
                    str(it.get("descricao", "") or ""),
                    str(it.get("unidade", "") or ""),
                    float(it.get("preco", 0) or 0),
                ))
        conn.commit()
    finally:
        conn.close()


def obter_tabela_precos(rede_id, fabrica_id):
    """Retorna lista de itens da tabela de preços para (rede_id, fabrica_id)."""
    rows = _fetch("""
        SELECT tp.codigo_fab, tp.descricao, tp.unidade,
               tp.preco, tp.data_atualizacao,
               p.nome as produto_nome
        FROM tabela_precos tp
        LEFT JOIN produtos p ON p.id = tp.produto_id
        WHERE tp.rede_id=%s AND tp.fabrica_id=%s
        ORDER BY tp.codigo_fab
    """, (rede_id, fabrica_id))
    return [dict(r) for r in rows]


def precos_tabela(rede_id, fabrica_id):
    """Retorna dict produto_id -> {preco: float, unidade: str} da tabela de preços."""
    rows = _fetch("""
        SELECT produto_id, preco, unidade
        FROM tabela_precos
        WHERE rede_id=%s AND fabrica_id=%s
          AND produto_id IS NOT NULL
    """, (rede_id, fabrica_id))
    return {
        r["produto_id"]: {
            "preco": float(r["preco"] or 0),
            "unidade": str(r["unidade"] or ""),
        }
        for r in rows
    }


def limpar_tabela_precos(rede_id, fabrica_id):
    _run(
        "DELETE FROM tabela_precos WHERE rede_id=%s AND fabrica_id=%s",
        (rede_id, fabrica_id)
    )


def vendas_por_loja_produto(fabrica_id, rede_id, data_inicio, data_fim):
    """Retorna dict: loja_id -> {produto_id -> {qtd, valor}}"""
    rows = _fetch("""
        SELECT f.loja_id, f.produto_id,
               SUM(f.qtd_vendida) as qtd,
               SUM(f.valor_total) as valor
        FROM faturamento f
        JOIN lojas l ON l.id = f.loja_id
        WHERE l.rede_id=%s
          AND f.produto_id IN (SELECT id FROM produtos WHERE fabrica_id=%s)
          AND (f.data_pedido = '' OR %s = '' OR f.data_pedido >= %s)
          AND (f.data_pedido = '' OR %s = '' OR f.data_pedido <= %s)
        GROUP BY f.loja_id, f.produto_id
    """, (rede_id, fabrica_id, data_inicio, data_inicio, data_fim, data_fim))
    result = {}
    for r in rows:
        lid = r["loja_id"]
        if lid not in result:
            result[lid] = {}
        result[lid][r["produto_id"]] = {
            "qtd": float(r["qtd"] or 0),
            "valor": float(r["valor"] or 0),
        }
    return result


# ── USUÁRIOS ──────────────────────────────────────────────────────────────────

def criar_usuario(username, nome, senha):
    """Cria usuário com senha criptografada. Retorna (True, '') ou (False, mensagem)."""
    pw_hash = bcrypt.hashpw(senha.encode(), bcrypt.gensalt()).decode()
    try:
        _run(
            "INSERT INTO usuarios (username, nome, password_hash, perfil) VALUES (%s, %s, %s, 'usuario')",
            (username.strip().lower(), nome.strip(), pw_hash)
        )
        return True, ""
    except Exception as e:
        msg = "Usuário já existe." if "unique" in str(e).lower() else str(e)
        return False, msg


def verificar_login(username, senha):
    """Verifica credenciais. Retorna dict do usuário ou None."""
    row = _fetch(
        "SELECT * FROM usuarios WHERE username=%s",
        (username.strip().lower(),), one=True
    )
    if not row:
        return None
    if bcrypt.checkpw(senha.encode(), row["password_hash"].encode()):
        _run(
            "UPDATE usuarios SET ultimo_acesso=NOW() WHERE username=%s",
            (username.strip().lower(),)
        )
        return dict(row)
    return None


def listar_usuarios():
    rows = _fetch("SELECT id, username, nome, perfil, ultimo_acesso FROM usuarios ORDER BY nome")
    return [dict(r) for r in rows]


def deletar_usuario(usuario_id):
    _run("DELETE FROM usuarios WHERE id=%s", (usuario_id,))


def alterar_senha(usuario_id, nova_senha):
    pw_hash = bcrypt.hashpw(nova_senha.encode(), bcrypt.gensalt()).decode()
    _run("UPDATE usuarios SET password_hash=%s WHERE id=%s", (pw_hash, usuario_id))
