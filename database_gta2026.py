import sqlite3, os, re
from pathlib import Path
from datetime import datetime

DB_PATH = Path(__file__).parent / "data" / "gta2026.db"

# ─────────────────────────────────────────────
# CONEXÃO
# ─────────────────────────────────────────────

def get_conn_2026():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


# ─────────────────────────────────────────────
# NORMALIZAÇÃO CPF/CNPJ
# ─────────────────────────────────────────────

def norm_cpf_2026(val):
    if not val:
        return ''
    s = str(val).strip()
    if re.match(r'^[\d,\.]+[Ee][+\-]?\d+$', s):
        try:
            s = str(int(float(s.replace(',', '.'))))
        except:
            return s
    s = re.sub(r'[^\d]', '', s)
    if not s:
        return ''
    s_sem_zero = s.lstrip('0') or '0'
    if len(s_sem_zero) <= 11:
        return s_sem_zero.zfill(11)
    else:
        return s_sem_zero.zfill(14)


# ─────────────────────────────────────────────
# INICIALIZAÇÃO DO BANCO
# ─────────────────────────────────────────────

def init_db_2026():
    """
    Garante que as tabelas de controle existem no gta2026.db.
    A tabela gta2026 já existe (criada pelo script de importação).
    """
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = get_conn_2026()
    c = conn.cursor()

    # Tabela principal (criada pelo script de importação, mas garantimos aqui)
    c.execute("""
        CREATE TABLE IF NOT EXISTS gta2026 (
            id TEXT, finalidade TEXT, data_emissao TEXT, situacao_pedido TEXT,
            gta_numero TEXT, total_animais TEXT, situacao_gta TEXT,
            data_hora_impressao TEXT, valor_dae TEXT, valor_fundepec TEXT,
            valor_total TEXT, origem_identificacao TEXT, origem_nome TEXT,
            origem_estabelecimento TEXT, origem_codigo_estabelecimento TEXT,
            origem_exploracao TEXT, origem_codigo_car TEXT,
            origem_estado_nome TEXT, origem_cidade_nome TEXT, origem_cidade_id TEXT,
            destinatario_identificacao TEXT, destinatario_nome TEXT,
            destinatario_estabelecimento TEXT, destinatario_codigo_estabelecimento TEXT,
            destinatario_codigo_car TEXT, destinatario_estado_nome TEXT,
            destinatario_cidade_nome TEXT, destinatario_cidade_id TEXT,
            transporte TEXT, criado_em TEXT, taxonomia_code TEXT,
            emitida_por_nome TEXT, taxonomia TEXT,
            bovino_macho_0_12 TEXT, bovino_macho_13_24 TEXT,
            bovino_macho_25_36 TEXT, bovino_macho_acima_36 TEXT,
            bovino_femea_0_12 TEXT, bovino_femea_13_24 TEXT,
            bovino_femea_25_36 TEXT, bovino_femea_acima_36 TEXT,
            bubalino_macho_0_12 TEXT, bubalino_macho_13_24 TEXT,
            bubalino_macho_25_36 TEXT, bubalino_macho_acima_36 TEXT,
            bubalino_femea_0_12 TEXT, bubalino_femea_13_24 TEXT,
            bubalino_femea_25_36 TEXT, bubalino_femea_acima_36 TEXT,
            caprino_macho_ate_12 TEXT, caprino_macho_acima_12 TEXT,
            caprino_femea_ate_12 TEXT, caprino_femea_acima_12 TEXT,
            ovino_macho_ate_12 TEXT, ovino_macho_acima_12 TEXT,
            ovino_femea_ate_12 TEXT, ovino_femea_acima_12 TEXT,
            equino_macho_ate_6 TEXT, equino_macho_acima_6 TEXT,
            equino_femea_ate_6 TEXT, equino_femea_acima_6 TEXT,
            muar_macho_ate_6 TEXT, muar_macho_acima_6 TEXT,
            muar_femea_ate_6 TEXT, muar_femea_acima_6 TEXT,
            asinino_macho_ate_6 TEXT, asinino_macho_acima_6 TEXT,
            asinino_femea_ate_6 TEXT, asinino_femea_acima_6 TEXT,
            galinha_ovos_ferteis TEXT, galinha_aves_1_dia TEXT,
            galinha_adulto TEXT, galinha_recriada TEXT,
            ganso_adulto TEXT, marreco_adulto TEXT,
            aves_nao_producao_adulto TEXT, suino_macho_reprodutor TEXT,
            suino_femea_matriz TEXT, suino_macho_leitao TEXT,
            suino_femea_leitao TEXT, suino_sexo_idade_nao_relevantes TEXT
        )
    """)

    # FTS5
    c.execute("""
        CREATE VIRTUAL TABLE IF NOT EXISTS gta2026_fts USING fts5(
            origem_nome,
            origem_identificacao,
            destinatario_nome,
            destinatario_identificacao,
            emitida_por_nome,
            gta_numero,
            origem_cidade_nome,
            destinatario_cidade_nome,
            content='gta2026',
            content_rowid='rowid'
        )
    """)

    # Controle de arquivos importados
    c.execute("""
        CREATE TABLE IF NOT EXISTS arquivos_importados_2026 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_arquivo TEXT NOT NULL UNIQUE,
            linhas INTEGER,
            importado_em TEXT NOT NULL
        )
    """)

    conn.commit()
    conn.close()


# ─────────────────────────────────────────────
# CONTROLE DE ARQUIVOS IMPORTADOS
# ─────────────────────────────────────────────

def arquivo_ja_importado_2026(nome):
    conn = get_conn_2026()
    r = conn.execute(
        "SELECT id FROM arquivos_importados_2026 WHERE nome_arquivo=?", (nome,)
    ).fetchone()
    conn.close()
    return r is not None


def registrar_arquivo_2026(nome, linhas):
    conn = get_conn_2026()
    conn.execute(
        "INSERT OR IGNORE INTO arquivos_importados_2026 (nome_arquivo, linhas, importado_em) VALUES (?,?,?)",
        (nome, linhas, datetime.now().isoformat())
    )
    conn.commit()
    conn.close()


def listar_arquivos_2026():
    conn = get_conn_2026()
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM arquivos_importados_2026 ORDER BY importado_em DESC"
    ).fetchall()]
    conn.close()
    return rows


# ─────────────────────────────────────────────
# IMPORTAÇÃO DE DATAFRAME
# ─────────────────────────────────────────────

MAPA_COLUNAS_2026 = {
    "id": "id",
    "finalidade": "finalidade",
    "data_emissao": "data_emissao",
    "situacao_pedido": "situacao_pedido",
    "gta_numero": "gta_numero",
    "total_animais": "total_animais",
    "situacao_gta": "situacao_gta",
    "data_hora_impressao": "data_hora_impressao",
    "valor_dae": "valor_dae",
    "valor_fundepec": "valor_fundepec",
    "valor_total": "valor_total",
    "origem_identificacao": "origem_identificacao",
    "origem_nome": "origem_nome",
    "origem_estabelecimento": "origem_estabelecimento",
    "origem_codigo_estabelecimento": "origem_codigo_estabelecimento",
    "origem_exploracao": "origem_exploracao",
    "origem_codigo_car": "origem_codigo_car",
    "origem_estado_nome": "origem_estado_nome",
    "origem_cidade_nome": "origem_cidade_nome",
    "origem_cidade_id": "origem_cidade_id",
    "destinatario_identificacao": "destinatario_identificacao",
    "destinatario_nome": "destinatario_nome",
    "destinatario_estabelecimento": "destinatario_estabelecimento",
    "destinatario_codigo_estabelecimento": "destinatario_codigo_estabelecimento",
    "destinatario_codigo_car": "destinatario_codigo_car",
    "destinatario_estado_nome": "destinatario_estado_nome",
    "destinatario_cidade_nome": "destinatario_cidade_nome",
    "destinatario_cidade_id": "destinatario_cidade_id",
    "transporte": "transporte",
    "criado_em": "criado_em",
    "taxonomia_code": "taxonomia_code",
    "emitida_por_nome": "emitida_por_nome",
    "taxonomia": "taxonomia",
    "BOVINO, MACHO, 0 A 12 MESES": "bovino_macho_0_12",
    "BOVINO, MACHO, 13 A 24 MESES": "bovino_macho_13_24",
    "BOVINO, MACHO, 25 A 36 MESES": "bovino_macho_25_36",
    "BOVINO, MACHO, ACIMA DE 36 MESES": "bovino_macho_acima_36",
    "BOVINO, FÊMEA, 0 A 12 MESES": "bovino_femea_0_12",
    "BOVINO, FÊMEA, 13 A 24 MESES": "bovino_femea_13_24",
    "BOVINO, FÊMEA, 25 A 36 MESES": "bovino_femea_25_36",
    "BOVINO, FÊMEA, ACIMA DE 36 MESES": "bovino_femea_acima_36",
    "BUBALINO, MACHO, 0 A 12 MESES": "bubalino_macho_0_12",
    "BUBALINO, MACHO, 13 A 24 MESES": "bubalino_macho_13_24",
    "BUBALINO, MACHO, 25 A 36 MESES": "bubalino_macho_25_36",
    "BUBALINO, MACHO, ACIMA DE 36 MESES": "bubalino_macho_acima_36",
    "BUBALINO, FÊMEA, 0 A 12 MESES": "bubalino_femea_0_12",
    "BUBALINO, FÊMEA, 13 A 24 MESES": "bubalino_femea_13_24",
    "BUBALINO, FÊMEA, 25 A 36 MESES": "bubalino_femea_25_36",
    "BUBALINO, FÊMEA, ACIMA DE 36 MESES": "bubalino_femea_acima_36",
    "CAPRINO, MACHO, ATÉ 12 MESES": "caprino_macho_ate_12",
    "CAPRINO, MACHO, ACIMA DE 12 MESES": "caprino_macho_acima_12",
    "CAPRINO, FÊMEA, ATÉ 12 MESES": "caprino_femea_ate_12",
    "CAPRINO, FÊMEA, ACIMA DE 12 MESES": "caprino_femea_acima_12",
    "OVINO, MACHO, ATÉ 12 MESES": "ovino_macho_ate_12",
    "OVINO, MACHO, ACIMA DE 12 MESES": "ovino_macho_acima_12",
    "OVINO, FÊMEA, ATÉ 12 MESES": "ovino_femea_ate_12",
    "OVINO, FÊMEA, ACIMA DE 12 MESES": "ovino_femea_acima_12",
    "EQUINO, MACHO, ATÉ 6 MESES": "equino_macho_ate_6",
    "EQUINO, MACHO, ACIMA DE 6 MESES": "equino_macho_acima_6",
    "EQUINO, FÊMEA, ATÉ 6 MESES": "equino_femea_ate_6",
    "EQUINO, FÊMEA, ACIMA DE 6 MESES": "equino_femea_acima_6",
    "MUAR, MACHO, ATÉ 6 MESES": "muar_macho_ate_6",
    "MUAR, MACHO, ACIMA DE 6 MESES": "muar_macho_acima_6",
    "MUAR, FÊMEA, ATÉ 6 MESES": "muar_femea_ate_6",
    "MUAR, FÊMEA, ACIMA DE 6 MESES": "muar_femea_acima_6",
    "ASININO, MACHO, ATÉ 6 MESES": "asinino_macho_ate_6",
    "ASININO, MACHO, ACIMA DE 6 MESES": "asinino_macho_acima_6",
    "ASININO, FÊMEA, ATÉ 6 MESES": "asinino_femea_ate_6",
    "ASININO, FÊMEA, ACIMA DE 6 MESES": "asinino_femea_acima_6",
    "GALINHA, OVOS FÉRTEIS": "galinha_ovos_ferteis",
    "GALINHA, AVES DE 1 DIA": "galinha_aves_1_dia",
    "GALINHA, ADULTO": "galinha_adulto",
    "GALINHA, RECRIADA": "galinha_recriada",
    "GANSO, ADULTO": "ganso_adulto",
    "MARRECO, ADULTO": "marreco_adulto",
    "AVES NÃO DESTINADAS À PRODUÇÃO, ADULTO": "aves_nao_producao_adulto",
    "SUÍNO, MACHO, REPRODUTOR (CACHAÇÃO)": "suino_macho_reprodutor",
    "SUÍNO, FÊMEA, MATRIZ": "suino_femea_matriz",
    "SUÍNO, MACHO, LEITÃO": "suino_macho_leitao",
    "SUÍNO, FÊMEA, LEITÃO": "suino_femea_leitao",
    "SUÍNO, SEXO E IDADE NÃO RELEVANTES": "suino_sexo_idade_nao_relevantes",
}

COLUNAS_DB_2026 = [
    "id", "finalidade", "data_emissao", "situacao_pedido", "gta_numero",
    "total_animais", "situacao_gta", "data_hora_impressao", "valor_dae",
    "valor_fundepec", "valor_total", "origem_identificacao", "origem_nome",
    "origem_estabelecimento", "origem_codigo_estabelecimento", "origem_exploracao",
    "origem_codigo_car", "origem_estado_nome", "origem_cidade_nome", "origem_cidade_id",
    "destinatario_identificacao", "destinatario_nome", "destinatario_estabelecimento",
    "destinatario_codigo_estabelecimento", "destinatario_codigo_car",
    "destinatario_estado_nome", "destinatario_cidade_nome", "destinatario_cidade_id",
    "transporte", "criado_em", "taxonomia_code", "emitida_por_nome", "taxonomia",
    "bovino_macho_0_12", "bovino_macho_13_24", "bovino_macho_25_36", "bovino_macho_acima_36",
    "bovino_femea_0_12", "bovino_femea_13_24", "bovino_femea_25_36", "bovino_femea_acima_36",
    "bubalino_macho_0_12", "bubalino_macho_13_24", "bubalino_macho_25_36", "bubalino_macho_acima_36",
    "bubalino_femea_0_12", "bubalino_femea_13_24", "bubalino_femea_25_36", "bubalino_femea_acima_36",
    "caprino_macho_ate_12", "caprino_macho_acima_12", "caprino_femea_ate_12", "caprino_femea_acima_12",
    "ovino_macho_ate_12", "ovino_macho_acima_12", "ovino_femea_ate_12", "ovino_femea_acima_12",
    "equino_macho_ate_6", "equino_macho_acima_6", "equino_femea_ate_6", "equino_femea_acima_6",
    "muar_macho_ate_6", "muar_macho_acima_6", "muar_femea_ate_6", "muar_femea_acima_6",
    "asinino_macho_ate_6", "asinino_macho_acima_6", "asinino_femea_ate_6", "asinino_femea_acima_6",
    "galinha_ovos_ferteis", "galinha_aves_1_dia", "galinha_adulto", "galinha_recriada",
    "ganso_adulto", "marreco_adulto", "aves_nao_producao_adulto",
    "suino_macho_reprodutor", "suino_femea_matriz", "suino_macho_leitao",
    "suino_femea_leitao", "suino_sexo_idade_nao_relevantes",
]


def importar_dataframe_2026(df, nome_arquivo):
    """
    Importa um DataFrame pandas para gta2026.db.
    Remove a segunda data_emissao duplicada automaticamente.
    """
    import pandas as pd

    # Remove segunda data_emissao duplicada
    indices_data = [i for i, c in enumerate(df.columns) if c == 'data_emissao']
    if len(indices_data) > 1:
        indices_remover = indices_data[1:]
        df = df.iloc[:, [i for i in range(len(df.columns)) if i not in indices_remover]]

    # Renomeia colunas usando o mapa
    novas_colunas = []
    ja_mapeou = {}
    for col in df.columns:
        col_strip = col.strip()
        if col_strip in MAPA_COLUNAS_2026:
            nome_db = MAPA_COLUNAS_2026[col_strip]
            if nome_db in ja_mapeou.values():
                nome_db = "_dup_" + col_strip[:20]
            ja_mapeou[col_strip] = nome_db
            novas_colunas.append(nome_db)
        else:
            novas_colunas.append("_ignorar_" + col_strip[:20])

    df.columns = novas_colunas

    # Remove colunas indesejadas
    cols_manter = [c for c in df.columns if not c.startswith("_ignorar_") and not c.startswith("_dup_")]
    df = df[cols_manter]

    # Garante todas as colunas do banco
    for col in COLUNAS_DB_2026:
        if col not in df.columns:
            df[col] = None

    df = df[COLUNAS_DB_2026]
    df = df.replace('', None)
    df = df.where(df.notna(), None)

    conn = get_conn_2026()
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")

    total = len(df)
    batch_size = 5000
    for start in range(0, total, batch_size):
        batch = df.iloc[start:start + batch_size]
        batch.to_sql('gta2026', conn, if_exists='append', index=False)

    # Rebuild FTS
    conn.execute("INSERT INTO gta2026_fts(gta2026_fts) VALUES('rebuild')")
    conn.commit()
    conn.close()

    registrar_arquivo_2026(nome_arquivo, total)
    return total


# ─────────────────────────────────────────────
# BUSCA
# ─────────────────────────────────────────────

def _fts_query_2026(nome):
    tokens = [t for t in nome.upper().split() if len(t) > 2]
    if not tokens:
        return None
    return ' AND '.join(f'"{t}"' for t in tokens)


def _nome_confere_2026(nome_pesquisado, campo):
    if not nome_pesquisado or not campo:
        return False
    tokens = [t for t in nome_pesquisado.upper().split() if len(t) > 2]
    campo_upper = str(campo).upper()
    matches = sum(1 for t in tokens if t in campo_upper)
    return matches >= min(2, len(tokens))


def buscar_gtas_2026(nome='', cpf='', emissor='', mes_ini=None, mes_fim=None):
    """
    Busca no banco gta2026.db.
    mes_ini / mes_fim: strings no formato 'YYYY-MM' para filtrar período.
    """
    conn = get_conn_2026()
    nome    = nome.strip().upper()
    cpf     = norm_cpf_2026(cpf)
    emissor = emissor.strip().upper()

    params = []
    sql_where = []

    # FTS
    fts_parts = []
    if nome:
        fts_q = _fts_query_2026(nome)
        if fts_q:
            fts_parts.append(f"origem_nome:{fts_q} OR destinatario_nome:{fts_q}")
    if emissor:
        fts_q_e = _fts_query_2026(emissor)
        if fts_q_e:
            fts_parts.append(f"emitida_por_nome:{fts_q_e}")

    if fts_parts:
        fts_full = ' OR '.join(fts_parts)
        sql_where.append("g.rowid IN (SELECT rowid FROM gta2026_fts WHERE gta2026_fts MATCH ?)")
        params.append(fts_full)

    if cpf:
        sql_where.append("(g.origem_identificacao LIKE ? OR g.destinatario_identificacao LIKE ?)")
        params += [f'%{cpf}%', f'%{cpf}%']

    if mes_ini:
        sql_where.append("g.data_emissao >= ?")
        params.append(mes_ini + '-01')

    if mes_fim:
        sql_where.append("g.data_emissao <= ?")
        params.append(mes_fim + '-31')

    if not sql_where:
        conn.close()
        return {}

    sql = f"SELECT g.* FROM gta2026 g WHERE {' AND '.join(sql_where)} ORDER BY g.data_emissao"
    rows = conn.execute(sql, params).fetchall()
    conn.close()

    resultado = {'2026': {'origem': [], 'destino': [], 'colunas': COLUNAS_DB_2026}}

    for row in rows:
        dados = dict(row)
        orig_id = norm_cpf_2026(dados.get('origem_identificacao') or '')
        dest_id = norm_cpf_2026(dados.get('destinatario_identificacao') or '')

        is_orig  = (cpf and cpf in orig_id) or (nome and _nome_confere_2026(nome, dados.get('origem_nome', '')))
        is_dest  = (cpf and cpf in dest_id) or (nome and _nome_confere_2026(nome, dados.get('destinatario_nome', '')))

        if is_dest and not is_orig:
            resultado['2026']['destino'].append(dados)
        else:
            resultado['2026']['origem'].append(dados)

    return resultado


def buscar_gtas_2026_lai():
    """
    Retorna todos os registros do banco para exportação LAI.
    """
    conn = get_conn_2026()
    rows = conn.execute(
        "SELECT * FROM gta2026 ORDER BY data_emissao"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def stats_2026():
    """Retorna estatísticas básicas do banco 2026."""
    conn = get_conn_2026()
    total = conn.execute("SELECT COUNT(*) FROM gta2026").fetchone()[0]
    datas = conn.execute("SELECT MIN(data_emissao), MAX(data_emissao) FROM gta2026").fetchone()
    arquivos = conn.execute("SELECT COUNT(*) FROM arquivos_importados_2026").fetchone()[0]
    conn.close()
    return {
        'total': total,
        'data_ini': datas[0],
        'data_fim': datas[1],
        'arquivos': arquivos,
    }
