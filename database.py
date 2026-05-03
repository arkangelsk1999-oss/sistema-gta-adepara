import sqlite3, os, re, json
from pathlib import Path

DB_PATH = Path(__file__).parent / "banco_gta.db"

# ── Detectar colunas de busca automaticamente ─────────────────
PADROES_ORIG_CPF  = ['cpf ou cnpj do produtor de origem', 'origem_identificacao']
PADROES_DEST_CPF  = ['cpf ou cnpj do produtor de destino', 'destinatario_identificacao']
PADROES_ORIG_NOME = ['nome do produtor de origem', 'origem_nome']
PADROES_DEST_NOME = ['nome do produtor de destino', 'destinatario_nome']

def detectar_col(colunas, padroes):
    cols_lower = {c.lower().strip(): c for c in colunas}
    for p in padroes:
        if p in cols_lower:
            return cols_lower[p]
    return None

def norm_cpf(val):
    if not val: return ''
    s = str(val).strip()
    if re.match(r'^[\d,\.]+[Ee][+\-]?\d+$', s):
        try: return str(int(float(s.replace(',','.'))))
        except: return s
    return re.sub(r'[^\d]', '', s)

def get_conn():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    c = conn.cursor()

    # Tabela de usuários
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        cpf TEXT NOT NULL UNIQUE,
        email TEXT NOT NULL UNIQUE,
        senha_hash TEXT NOT NULL,
        orgao TEXT NOT NULL,
        nivel TEXT NOT NULL DEFAULT 'usuario',
        ativo INTEGER NOT NULL DEFAULT 1,
        criado_em TEXT NOT NULL,
        ultimo_acesso TEXT
    )''')

    # Tabela de GTAs — linha completa como JSON + campos indexados
    c.execute('''CREATE TABLE IF NOT EXISTS gtas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ano INTEGER NOT NULL,
        arquivo TEXT NOT NULL,
        orig_cpf TEXT,
        orig_nome TEXT,
        dest_cpf TEXT,
        dest_nome TEXT,
        dados_json TEXT NOT NULL
    )''')

    # Índices de busca
    c.execute('CREATE INDEX IF NOT EXISTS idx_orig_cpf  ON gtas(orig_cpf)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_dest_cpf  ON gtas(dest_cpf)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_orig_nome ON gtas(orig_nome)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_dest_nome ON gtas(dest_nome)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_ano       ON gtas(ano)')

    # Tabela de arquivos já importados
    c.execute('''CREATE TABLE IF NOT EXISTS arquivos_importados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome_arquivo TEXT NOT NULL UNIQUE,
        ano INTEGER,
        linhas INTEGER,
        importado_em TEXT NOT NULL
    )''')

    # Tabela de auditoria
    c.execute('''CREATE TABLE IF NOT EXISTS auditoria (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data_hora TEXT NOT NULL,
        usuario_id INTEGER,
        usuario_nome TEXT,
        usuario_login TEXT,
        orgao TEXT,
        ip TEXT,
        localidade TEXT,
        cpf_cnpj_pesquisado TEXT,
        nome_pesquisado TEXT,
        total_resultados INTEGER
    )''')

    # Founder padrão — criado só se não existir
    from werkzeug.security import generate_password_hash
    from datetime import datetime
    existe = c.execute("SELECT id FROM usuarios WHERE nivel='founder'").fetchone()
    if not existe:
        c.execute('''INSERT INTO usuarios 
            (nome, cpf, email, senha_hash, orgao, nivel, ativo, criado_em)
            VALUES (?,?,?,?,?,?,?,?)''', (
            'Gilliard Costa Rodrigues',
            '00000000000',  # CPF real configurado no primeiro acesso
            'founder@arkangelsk.com',
            generate_password_hash('Arkangelsk@2025'),
            'Arkangelsk',
            'founder',
            1,
            datetime.now().isoformat()
        ))

    conn.commit()
    conn.close()

def arquivo_ja_importado(nome):
    conn = get_conn()
    r = conn.execute("SELECT id FROM arquivos_importados WHERE nome_arquivo=?", (nome,)).fetchone()
    conn.close()
    return r is not None

def registrar_arquivo(nome, ano, linhas):
    from datetime import datetime
    conn = get_conn()
    conn.execute(
        "INSERT OR IGNORE INTO arquivos_importados (nome_arquivo,ano,linhas,importado_em) VALUES (?,?,?,?)",
        (nome, ano, linhas, datetime.now().isoformat())
    )
    conn.commit()
    conn.close()

def importar_dataframe(df, ano, nome_arquivo):
    import json
    cols = list(df.columns)
    col_orig_cpf  = detectar_col(cols, PADROES_ORIG_CPF)
    col_dest_cpf  = detectar_col(cols, PADROES_DEST_CPF)
    col_orig_nome = detectar_col(cols, PADROES_ORIG_NOME)
    col_dest_nome = detectar_col(cols, PADROES_DEST_NOME)

    conn = get_conn()
    c = conn.cursor()
    lote = []

    for _, row in df.iterrows():
        orig_cpf  = norm_cpf(row.get(col_orig_cpf,  '')) if col_orig_cpf  else ''
        dest_cpf  = norm_cpf(row.get(col_dest_cpf,  '')) if col_dest_cpf  else ''
        orig_nome = str(row.get(col_orig_nome, '')).upper().strip() if col_orig_nome else ''
        dest_nome = str(row.get(col_dest_nome, '')).upper().strip() if col_dest_nome else ''
        dados     = json.dumps(row.to_dict(), ensure_ascii=False, default=str)

        lote.append((ano, nome_arquivo, orig_cpf, orig_nome, dest_cpf, dest_nome, dados))

        if len(lote) >= 5000:
            c.executemany(
                "INSERT INTO gtas (ano,arquivo,orig_cpf,orig_nome,dest_cpf,dest_nome,dados_json) VALUES (?,?,?,?,?,?,?)",
                lote
            )
            conn.commit()
            lote = []

    if lote:
        c.executemany(
            "INSERT INTO gtas (ano,arquivo,orig_cpf,orig_nome,dest_cpf,dest_nome,dados_json) VALUES (?,?,?,?,?,?,?)",
            lote
        )
        conn.commit()
    conn.close()
    return len(df)

def buscar_gtas(nome='', cpf='', ano_ini=None, ano_fim=None):
    import json
    conn = get_conn()
    nome = nome.strip().upper()
    cpf  = norm_cpf(cpf)

    condicoes = []
    params = []

    if nome and cpf:
        condicoes.append("(orig_nome LIKE ? OR dest_nome LIKE ? OR orig_cpf LIKE ? OR dest_cpf LIKE ?)")
        params += [f'%{nome}%', f'%{nome}%', f'%{cpf}%', f'%{cpf}%']
    elif nome:
        condicoes.append("(orig_nome LIKE ? OR dest_nome LIKE ?)")
        params += [f'%{nome}%', f'%{nome}%']
    elif cpf:
        condicoes.append("(orig_cpf LIKE ? OR dest_cpf LIKE ?)")
        params += [f'%{cpf}%', f'%{cpf}%']
    else:
        return {}

    if ano_ini:
        condicoes.append("ano >= ?")
        params.append(int(ano_ini))
    if ano_fim:
        condicoes.append("ano <= ?")
        params.append(int(ano_fim))

    sql = f"SELECT * FROM gtas WHERE {' AND '.join(condicoes)} ORDER BY ano"
    rows = conn.execute(sql, params).fetchall()
    conn.close()

    # Organizar por ano, separando origem e destino
    resultado = {}
    for row in rows:
        ano  = row['ano']
        dados = json.loads(row['dados_json'])
        # Determinar papel
        is_orig = (nome and nome in (row['orig_nome'] or '')) or (cpf and cpf in (row['orig_cpf'] or ''))
        is_dest = (nome and nome in (row['dest_nome'] or '')) or (cpf and cpf in (row['dest_cpf'] or ''))

        if ano not in resultado:
            resultado[ano] = {'origem': [], 'destino': [], 'colunas': list(dados.keys())}
        if is_orig:
            resultado[ano]['origem'].append(dados)
        if is_dest and not is_orig:
            resultado[ano]['destino'].append(dados)

    return resultado

def registrar_auditoria(usuario, ip, localidade, cpf_pesquisado, nome_pesquisado, total):
    from datetime import datetime
    conn = get_conn()
    conn.execute('''INSERT INTO auditoria 
        (data_hora,usuario_id,usuario_nome,usuario_login,orgao,ip,localidade,cpf_cnpj_pesquisado,nome_pesquisado,total_resultados)
        VALUES (?,?,?,?,?,?,?,?,?,?)''', (
        datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
        usuario.get('id'),
        usuario.get('nome'),
        usuario.get('email'),
        usuario.get('orgao'),
        ip,
        localidade,
        cpf_pesquisado,
        nome_pesquisado,
        total
    ))
    conn.commit()
    conn.close()
