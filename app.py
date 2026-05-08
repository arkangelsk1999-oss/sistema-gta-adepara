from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path
import pandas as pd
import json, os, re, io, threading

from database import (
    init_db, get_conn, buscar_gtas, importar_dataframe,
    arquivo_ja_importado, registrar_arquivo, registrar_auditoria, norm_cpf,
    verificar_aceite_termos, registrar_aceite_termos,
    resolver_localidade, listar_ip_localidade, salvar_ip_localidade, excluir_ip_localidade
)
from relatorio import gerar_excel_resultado, gerar_pdf_auditoria, gerar_csv_resultado

app = Flask(__name__, 
            template_folder='.', 
            static_folder='.',
            static_url_path='')
app.secret_key = os.environ.get('SECRET_KEY', 'arkangelsk-2025-chave-secreta')

UPLOAD_FOLDER = Path(__file__).parent / 'uploads_temp'
UPLOAD_FOLDER.mkdir(exist_ok=True)
ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}

# ── Controle de tentativas de login ───────────────────────────
tentativas_login = {}
MAX_TENTATIVAS   = 5
BLOQUEIO_MINUTOS = 15
INATIVIDADE_MINUTOS = 30

def verificar_bloqueio(ip):
    if ip not in tentativas_login:
        return False, 0
    dados = tentativas_login[ip]
    if 'bloqueado_ate' in dados:
        restante = (dados['bloqueado_ate'] - datetime.now()).total_seconds()
        if restante > 0:
            return True, int(restante)
        else:
            del tentativas_login[ip]
    return False, 0

def registrar_tentativa(ip):
    if ip not in tentativas_login:
        tentativas_login[ip] = {'tentativas': 0}
    tentativas_login[ip]['tentativas'] += 1
    if tentativas_login[ip]['tentativas'] >= MAX_TENTATIVAS:
        tentativas_login[ip]['bloqueado_ate'] = datetime.now() + timedelta(minutes=BLOQUEIO_MINUTOS)
        return True
    return False

def limpar_tentativas(ip):
    if ip in tentativas_login:
        del tentativas_login[ip]

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'usuario_id' not in session:
            return redirect(url_for('login'))
        ultimo = session.get('ultimo_acesso')
        if ultimo:
            ultimo_dt = datetime.fromisoformat(ultimo)
            if datetime.now() - ultimo_dt > timedelta(minutes=INATIVIDADE_MINUTOS):
                session.clear()
                return redirect(url_for('login', expirado=1))
        if request.endpoint not in ['exportar_excel', 'exportar_csv']:
            session['ultimo_acesso'] = datetime.now().isoformat()
        return f(*args, **kwargs)
    return decorated

def nivel_required(*niveis):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if session.get('nivel') not in niveis:
                return redirect(url_for('index'))
            return f(*args, **kwargs)
        return decorated
    return decorator

def get_usuario_session():
    return {
        'id':    session.get('usuario_id'),
        'nome':  session.get('usuario_nome'),
        'email': session.get('usuario_email'),
        'orgao': session.get('usuario_orgao'),
        'nivel': session.get('nivel'),
        'cpf':   session.get('usuario_cpf'),
    }

def extrair_ano(nome):
    m = re.search(r'20\d{2}', nome)
    return int(m.group()) if m else 9999

@app.route('/login', methods=['GET','POST'])
def login():
    ip = request.headers.get('X-Forwarded-For', request.remote_addr)
    erro = None
    expirado = request.args.get('expirado')

    bloqueado, segundos = verificar_bloqueio(ip)
    if bloqueado:
        minutos = segundos // 60
        segs    = segundos % 60
        erro = f"Acesso bloqueado por tentativas excessivas. Tente novamente em {minutos}m {segs}s."
        return render_template('login.html', erro=erro, bloqueado=True, segundos=segundos)

    if request.method == 'POST':
        email = request.form.get('email','').strip().lower()
        senha = request.form.get('senha','')
        conn  = get_conn()
        user  = conn.execute(
            "SELECT * FROM usuarios WHERE LOWER(email)=? AND ativo=1", (email,)
        ).fetchone()
        conn.close()

        if user and check_password_hash(user['senha_hash'], senha):
            limpar_tentativas(ip)
            session['usuario_id']    = user['id']
            session['usuario_nome']  = user['nome']
            session['usuario_email'] = user['email']
            session['usuario_orgao'] = user['orgao']
            session['usuario_cpf']   = user['cpf']
            session['nivel']         = user['nivel']
            session['ultimo_acesso'] = datetime.now().isoformat()
            conn = get_conn()
            conn.execute("UPDATE usuarios SET ultimo_acesso=? WHERE id=?",
                         (datetime.now().isoformat(), user['id']))
            conn.commit()
            conn.close()

            if not verificar_aceite_termos(user['id']):
                return redirect(url_for('termos'))

            return redirect(url_for('index'))
        else:
            bloqueou = registrar_tentativa(ip)
            dados = tentativas_login.get(ip, {})
            tentativas = dados.get('tentativas', 0)
            restantes = MAX_TENTATIVAS - tentativas
            if bloqueou:
                erro = f"Acesso bloqueado por {BLOQUEIO_MINUTOS} minutos devido a tentativas excessivas."
            else:
                erro = f"E-mail ou senha incorretos. {restantes} tentativa(s) restante(s) antes do bloqueio."

    return render_template('login.html', erro=erro, bloqueado=False, segundos=0, expirado=expirado)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/termos')
def termos():
    if 'usuario_id' not in session:
        return redirect(url_for('login'))
    conn = get_conn()
    versao_ativa = conn.execute(
        "SELECT versao FROM termos_versao WHERE ativo=1 ORDER BY id DESC LIMIT 1"
    ).fetchone()
    conn.close()
    versao = versao_ativa['versao'] if versao_ativa else '1.0'
    return render_template('termos.html', versao=versao)

@app.route('/termos/aceitar', methods=['POST'])
def termos_aceitar():
    if 'usuario_id' not in session:
        return redirect(url_for('login'))
    ip = request.headers.get('X-Forwarded-For', request.remote_addr)
    registrar_aceite_termos(
        usuario_id=session['usuario_id'],
        usuario_nome=session['usuario_nome'],
        usuario_cpf=session.get('usuario_cpf', ''),
        ip=ip
    )
    return redirect(url_for('index'))

@app.route('/termos/recusar', methods=['POST'])
def termos_recusar():
    session.clear()
    return redirect(url_for('login'))

@app.route('/ping')
@login_required
def ping():
    return jsonify({'ok': True})

@app.route('/')
@login_required
def index():
    return render_template('index.html', usuario=get_usuario_session())

@app.route('/buscar', methods=['POST'])
@login_required
def buscar():
    nome    = request.form.get('nome', '').strip()
    cpf     = request.form.get('cpf', '').strip()
    emissor = request.form.get('emissor', '').strip()
    ano_ini = request.form.get('ano_ini', '').strip()
    ano_fim = request.form.get('ano_fim', '').strip()

    if not nome and not cpf and not emissor:
        return jsonify({'erro': 'Informe nome, CPF/CNPJ ou usuário emissor'}), 400

    try:
        resultado = buscar_gtas(nome=nome, cpf=cpf, emissor=emissor,
                                ano_ini=ano_ini or None,
                                ano_fim=ano_fim or None)

        total = sum(len(v['origem']) + len(v['destino']) for v in resultado.values())

        ip = request.headers.get('X-Forwarded-For', request.remote_addr)
        localidade = resolver_localidade(ip)
        registrar_auditoria(
            usuario=get_usuario_session(),
            ip=ip,
            localidade=localidade,
            cpf_pesquisado=norm_cpf(cpf),
            nome_pesquisado=nome.upper() or emissor.upper(),
            total=total
        )

        if not resultado:
            return jsonify({'vazio': True, 'total': 0})

        anos = sorted(resultado.keys())
        resumo = {
            str(ano): {
                'origem':  len(resultado[ano]['origem']),
                'destino': len(resultado[ano]['destino']),
            }
            for ano in anos
        }

        session['ultimo_resultado'] = json.dumps({
            'nome': nome, 'cpf': cpf, 'emissor': emissor,
        })

        preview = {}
        for ano in anos:
            orig = resultado[ano]['origem'][:10]
            dest = resultado[ano]['destino'][:10]
            cols = resultado[ano]['colunas']
            preview[str(ano)] = {'colunas': cols, 'origem': orig, 'destino': dest}

        return jsonify({
            'total':   total,
            'anos':    [str(a) for a in anos],
            'resumo':  resumo,
            'preview': preview,
            'nome_encontrado': nome or cpf or emissor,
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'erro': str(e)}), 500

@app.route('/exportar/excel', methods=['POST'])
@login_required
def exportar_excel():
    nome    = request.form.get('nome', '')
    cpf     = request.form.get('cpf', '')
    emissor = request.form.get('emissor', '')
    ano_ini = request.form.get('ano_ini', '')
    ano_fim = request.form.get('ano_fim', '')

    resultado = buscar_gtas(nome=nome, cpf=cpf, emissor=emissor,
                             ano_ini=ano_ini or None, ano_fim=ano_fim or None)
    if not resultado:
        return 'Sem resultados', 404

    buf = gerar_excel_resultado(resultado, nome or emissor, cpf, usuario=get_usuario_session())
    nome_arquivo = f"GTA_{(nome or cpf or emissor).replace(' ','_')[:30]}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=nome_arquivo,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/exportar/csv', methods=['POST'])
@login_required
@nivel_required('founder')
def exportar_csv():
    nome    = request.form.get('nome', '')
    cpf     = request.form.get('cpf', '')
    emissor = request.form.get('emissor', '')
    ano_ini = request.form.get('ano_ini', '')
    ano_fim = request.form.get('ano_fim', '')

    resultado = buscar_gtas(nome=nome, cpf=cpf, emissor=emissor,
                             ano_ini=ano_ini or None, ano_fim=ano_fim or None)
    if not resultado:
        return 'Sem resultados', 404

    buf = gerar_csv_resultado(resultado, nome or emissor, cpf, usuario=get_usuario_session())
    nome_arquivo = f"GTA_{(nome or cpf or emissor).replace(' ','_')[:30]}_{datetime.now().strftime('%Y%m%d')}.csv"
    return send_file(buf, as_attachment=True, download_name=nome_arquivo, mimetype='text/csv')

@app.route('/auditoria')
@login_required
@nivel_required('founder', 'master')
def auditoria():
    return render_template('auditoria.html', usuario=get_usuario_session())

@app.route('/auditoria/dados')
@login_required
@nivel_required('founder', 'master')
def auditoria_dados():
    data_ini = request.args.get('data_ini', '')
    data_fim = request.args.get('data_fim', '')
    usuario  = request.args.get('usuario', '')
    conn = get_conn()
    sql  = "SELECT * FROM auditoria WHERE 1=1"
    params = []
    if data_ini:
        sql += " AND data_hora >= ?"; params.append(data_ini)
    if data_fim:
        sql += " AND data_hora <= ?"; params.append(data_fim + ' 23:59:59')
    if usuario:
        sql += " AND (usuario_nome LIKE ? OR usuario_login LIKE ?)"; params += [f'%{usuario}%', f'%{usuario}%']
    sql += " ORDER BY data_hora DESC LIMIT 1000"
    rows = [dict(r) for r in conn.execute(sql, params).fetchall()]
    conn.close()
    return jsonify(rows)

@app.route('/auditoria/pdf')
@login_required
@nivel_required('founder', 'master')
def auditoria_pdf():
    data_ini = request.args.get('data_ini', '')
    data_fim = request.args.get('data_fim', '')
    usuario  = request.args.get('usuario', '')
    conn = get_conn()
    sql  = "SELECT * FROM auditoria WHERE 1=1"
    params = []
    if data_ini:
        sql += " AND data_hora >= ?"; params.append(data_ini)
    if data_fim:
        sql += " AND data_hora <= ?"; params.append(data_fim + ' 23:59:59')
    if usuario:
        sql += " AND (usuario_nome LIKE ? OR usuario_login LIKE ?)"; params += [f'%{usuario}%', f'%{usuario}%']
    sql += " ORDER BY data_hora DESC"
    rows = [dict(r) for r in conn.execute(sql, params).fetchall()]
    conn.close()
    gerado_por = get_usuario_session()
    buf = gerar_pdf_auditoria(rows, gerado_por, data_ini, data_fim)
    return send_file(buf, as_attachment=True,
                     download_name=f"auditoria_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                     mimetype='application/pdf')

@app.route('/usuarios')
@login_required
@nivel_required('founder', 'master')
def usuarios():
    conn = get_conn()
    lista = [dict(r) for r in conn.execute(
        "SELECT id,nome,cpf,email,orgao,nivel,ativo,criado_em,ultimo_acesso FROM usuarios WHERE nivel != 'founder' ORDER BY nome"
    ).fetchall()]
    conn.close()
    return render_template('usuarios.html', usuarios=lista, usuario=get_usuario_session())

@app.route('/usuarios/novo', methods=['POST'])
@login_required
@nivel_required('founder', 'master')
def novo_usuario():
    dados = request.get_json()
    nome  = dados.get('nome','').strip()
    cpf   = norm_cpf(dados.get('cpf',''))
    email = dados.get('email','').strip().lower()
    orgao = dados.get('orgao','').strip()
    nivel = dados.get('nivel','usuario')
    senha = dados.get('senha','')

    if not all([nome, cpf, email, orgao, senha]):
        return jsonify({'erro': 'Preencha todos os campos'}), 400
    if nivel == 'founder' and session.get('nivel') != 'founder':
        return jsonify({'erro': 'Sem permissão'}), 403

    try:
        conn = get_conn()
        conn.execute('''INSERT INTO usuarios (nome,cpf,email,senha_hash,orgao,nivel,ativo,criado_em)
                        VALUES (?,?,?,?,?,?,1,?)''',
                     (nome, cpf, email, generate_password_hash(senha), orgao, nivel,
                      datetime.now().isoformat()))
        conn.commit()
        conn.close()
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'erro': str(e)}), 400

@app.route('/usuarios/<int:uid>/toggle', methods=['POST'])
@login_required
@nivel_required('founder', 'master')
def toggle_usuario(uid):
    conn = get_conn()
    user = conn.execute("SELECT * FROM usuarios WHERE id=?", (uid,)).fetchone()
    if not user:
        conn.close()
        return jsonify({'erro': 'Não encontrado'}), 404
    novo = 0 if user['ativo'] else 1
    conn.execute("UPDATE usuarios SET ativo=? WHERE id=?", (novo, uid))
    conn.commit()
    conn.close()
    return jsonify({'ativo': novo})

@app.route('/importar')
@login_required
@nivel_required('founder', 'master')
def importar_page():
    conn = get_conn()
    importados = [dict(r) for r in conn.execute(
        "SELECT * FROM arquivos_importados ORDER BY ano DESC"
    ).fetchall()]
    conn.close()
    return render_template('importar.html', importados=importados, usuario=get_usuario_session())

status_importacao = {}

@app.route('/importar/upload', methods=['POST'])
@login_required
@nivel_required('founder', 'master')
def upload_arquivo():
    if 'arquivo' not in request.files:
        return jsonify({'erro': 'Nenhum arquivo enviado'}), 400

    f    = request.files['arquivo']
    nome = secure_filename(f.name if hasattr(f,'name') else f.filename)
    nome = f.filename
    ext  = nome.rsplit('.', 1)[-1].lower()

    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({'erro': 'Formato não suportado. Use CSV ou XLSX'}), 400
    if arquivo_ja_importado(nome):
        return jsonify({'erro': f'Arquivo "{nome}" já foi importado anteriormente'}), 400

    caminho = UPLOAD_FOLDER / nome
    f.save(str(caminho))

    job_id = nome
    status_importacao[job_id] = {'status': 'processando', 'progresso': 0, 'msg': 'Iniciando...'}

    def processar():
        try:
            ano = extrair_ano(nome)
            status_importacao[job_id]['msg'] = 'Lendo arquivo...'
            if ext == 'csv':
                df = pd.read_csv(str(caminho), encoding='latin-1', sep=';', dtype=str, low_memory=False)
            else:
                abas = pd.read_excel(str(caminho), sheet_name=None, dtype=str)
                df   = pd.concat(abas.values(), ignore_index=True)

            total = len(df)
            status_importacao[job_id]['msg'] = f'Importando {total:,} linhas...'
            status_importacao[job_id]['progresso'] = 30

            linhas = importar_dataframe(df, ano, nome)
            registrar_arquivo(nome, ano, linhas)

            status_importacao[job_id] = {
                'status': 'concluido', 'progresso': 100,
                'msg': f'✅ {linhas:,} linhas importadas com sucesso!'
            }
        except Exception as e:
            status_importacao[job_id] = {
                'status': 'erro', 'progresso': 0, 'msg': f'❌ Erro: {str(e)}'
            }
        finally:
            if caminho.exists():
                caminho.unlink()

    threading.Thread(target=processar, daemon=True).start()
    return jsonify({'job_id': job_id, 'ok': True})

@app.route('/importar/status/<job_id>')
@login_required
def status_job(job_id):
    return jsonify(status_importacao.get(job_id, {'status': 'desconhecido'}))

# ── Alteração de senha ────────────────────────────────────────
@app.route('/alterar-senha', methods=['GET', 'POST'])
@login_required
def alterar_senha():
    erro = None
    sucesso = None

    if request.method == 'POST':
        senha_atual = request.form.get('senha_atual', '')
        senha_nova  = request.form.get('senha_nova', '').strip()
        senha_conf  = request.form.get('senha_conf', '').strip()

        conn = get_conn()
        user = conn.execute("SELECT * FROM usuarios WHERE id=?",
                            (session['usuario_id'],)).fetchone()
        conn.close()

        if not check_password_hash(user['senha_hash'], senha_atual):
            erro = "Senha atual incorreta."
        elif senha_nova != senha_conf:
            erro = "A nova senha e a confirmação não coincidem."
        elif len(senha_nova) < 8:
            erro = "A nova senha deve ter ao menos 8 caracteres."
        elif not re.search(r'[A-Z]', senha_nova):
            erro = "A nova senha deve conter ao menos uma letra maiúscula."
        elif not re.search(r'[0-9]', senha_nova):
            erro = "A nova senha deve conter ao menos um número."
        elif not re.search(r'[^A-Za-z0-9]', senha_nova):
            erro = "A nova senha deve conter ao menos um caractere especial (@, #, !, $, etc.)."
        else:
            conn = get_conn()
            conn.execute("UPDATE usuarios SET senha_hash=? WHERE id=?",
                         (generate_password_hash(senha_nova), session['usuario_id']))
            conn.commit()
            conn.close()
            sucesso = "Senha alterada com sucesso!"

    return render_template('alterar_senha.html',
                           usuario=get_usuario_session(),
                           erro=erro, sucesso=sucesso)

# ── Download do manual ────────────────────────────────────────
@app.route('/manual')
def manual_pdf():
    caminho = Path(__file__).parent / 'Manual_e_Termos_SistemaGTA_ADEPARA.pdf'
    return send_file(caminho, as_attachment=True,
                     download_name='Manual_SistemaGTA_ADEPARA.pdf',
                     mimetype='application/pdf')

# ── Gestão de IP/Localidade ───────────────────────────────────
@app.route('/ip-localidade')
@login_required
@nivel_required('founder', 'master')
def ip_localidade():
    lista = listar_ip_localidade()
    return render_template('ip_localidade.html', lista=lista, usuario=get_usuario_session())

@app.route('/ip-localidade/salvar', methods=['POST'])
@login_required
@nivel_required('founder', 'master')
def ip_localidade_salvar():
    dados = request.get_json()
    ip         = dados.get('ip', '').strip()
    localidade = dados.get('localidade', '').strip()
    descricao  = dados.get('descricao', '').strip()
    if not ip or not localidade:
        return jsonify({'erro': 'IP e Localidade são obrigatórios'}), 400
    try:
        salvar_ip_localidade(ip, localidade, descricao)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'erro': str(e)}), 400

@app.route('/ip-localidade/excluir/<int:id>', methods=['POST'])
@login_required
@nivel_required('founder', 'master')
def ip_localidade_excluir(id):
    excluir_ip_localidade(id)
    return jsonify({'ok': True})

init_db()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)