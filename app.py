"""
O2 Inc - Web App de Auditoria de Lançamentos Omie
"""

import os
import uuid
from dotenv import load_dotenv
load_dotenv()
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename

from audit_engine import executar_auditoria
from excel_generator import gerar_excel
from database import (
    init_db, salvar_relatorio, listar_empresas,
    listar_relatorios_empresa, buscar_relatorio,
    REPORTS_DIR,
)

app = Flask(__name__)
IS_VERCEL = os.getenv('VERCEL') == '1'
# Vercel Serverless has stricter body/runtime limits than local/VM deploys.
app.config['MAX_CONTENT_LENGTH'] = (4 if IS_VERCEL else 50) * 1024 * 1024

UPLOAD_FOLDER = '/tmp/o2_uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

init_db()

EXTENSOES_PERMITIDAS = {'xlsx', 'xls'}


def permitido(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in EXTENSOES_PERMITIDAS


@app.errorhandler(413)
def arquivo_muito_grande(_err):
    limite_mb = app.config['MAX_CONTENT_LENGTH'] // (1024 * 1024)
    return jsonify({
        'erro': f'Arquivo excede o limite de {limite_mb}MB para este ambiente.'
    }), 413


# ── Páginas ──────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/historico')
def historico():
    return render_template('historico.html')


@app.route('/healthz')
def healthz():
    return jsonify({
        'ok': True,
        'service': 'auditoria-o2',
        'version': '2026-05-25-competencia-registro-global-v4',
    })


# ── Auditoria ─────────────────────────────────────────────────────────────────

@app.route('/auditar', methods=['POST'])
def auditar():
    if 'arquivo' not in request.files:
        return jsonify({'erro': 'Nenhum arquivo enviado.'}), 400

    arquivo = request.files['arquivo']
    empresa = request.form.get('empresa', '').strip()

    if not arquivo.filename or not permitido(arquivo.filename):
        return jsonify({'erro': 'Formato inválido. Envie um arquivo .xlsx ou .xls do Omie.'}), 400

    if not empresa:
        return jsonify({'erro': 'Informe o nome da empresa antes de auditar.'}), 400

    nome_original = secure_filename(arquivo.filename)
    session_id    = str(uuid.uuid4())[:8]
    caminho_upload = os.path.join(UPLOAD_FOLDER, f'{session_id}_{nome_original}')
    arquivo.save(caminho_upload)

    try:
        resultado = executar_auditoria(
            caminho_upload,
            usar_registro_como_competencia=True,
        )
    except Exception as e:
        os.remove(caminho_upload)
        return jsonify({'erro': f'Erro ao processar arquivo: {str(e)}'}), 500

    nome_excel    = f'auditoria_o2_{session_id}.xlsx'
    caminho_excel = os.path.join(REPORTS_DIR, nome_excel)
    gerar_excel(resultado, caminho_excel, nome_original, empresa, source_path=caminho_upload)
    os.remove(caminho_upload)

    rid = salvar_relatorio(
        empresa             = empresa,
        data_analise        = resultado['data_analise'],
        nome_arquivo_origem = nome_original,
        total_lancamentos   = resultado['total_lancamentos'],
        total_erros         = resultado['total_erros'],
        por_severidade      = resultado['por_severidade'],
        por_regra           = resultado['por_regra'],
        caminho_arquivo     = caminho_excel,
    )

    return jsonify({
        'sucesso':           True,
        'relatorio_id':      rid,
        'total_lancamentos': resultado['total_lancamentos'],
        'total_erros':       resultado['total_erros'],
        'por_severidade':    resultado['por_severidade'],
        'por_regra':         resultado['por_regra'],
        'download_url':      f'/download/relatorio/{rid}',
        # Amostra para depuração rápida no frontend/API sem inflar demais o payload.
        'erros_debug':       resultado.get('erros', [])[:100],
    })


# ── Download ──────────────────────────────────────────────────────────────────

@app.route('/download/relatorio/<int:rid>')
def download_relatorio(rid):
    r = buscar_relatorio(rid)
    if not r:
        return 'Relatório não encontrado.', 404
    caminho = r['caminho_arquivo']
    if not os.path.exists(caminho):
        return 'Arquivo não encontrado no servidor.', 404
    empresa_safe = secure_filename(r['empresa']) or 'relatorio'
    data_safe    = r['data_analise'].replace('/', '-').replace(':', '-').replace(' ', '_')[:16]
    return send_file(
        caminho,
        as_attachment=True,
        download_name=f'auditoria_{empresa_safe}_{data_safe}.xlsx',
    )


# ── API histórico ─────────────────────────────────────────────────────────────

@app.route('/api/empresas')
def api_empresas():
    return jsonify(listar_empresas())


@app.route('/api/relatorios')
def api_relatorios():
    empresa = request.args.get('empresa', '').strip()
    if not empresa:
        return jsonify([])
    relatorios = listar_relatorios_empresa(empresa)
    for r in relatorios:
        r['download_url'] = f'/download/relatorio/{r["id"]}'
        r.pop('caminho_arquivo', None)
    return jsonify(relatorios)


@app.route('/api/comparar')
def api_comparar():
    try:
        id1 = int(request.args.get('id1', 0))
        id2 = int(request.args.get('id2', 0))
    except ValueError:
        return jsonify({'erro': 'IDs inválidos'}), 400

    r1 = buscar_relatorio(id1)
    r2 = buscar_relatorio(id2)
    if not r1 or not r2:
        return jsonify({'erro': 'Relatório não encontrado'}), 404

    def delta(a, b):
        d = b - a
        pct = (d / a * 100) if a != 0 else (100.0 if b > 0 else 0.0)
        return {'a': a, 'b': b, 'delta': d, 'delta_pct': round(pct, 1)}

    sev_comp = {}
    for k in ['CRÍTICO', 'ALTO', 'MÉDIO', 'BAIXO']:
        sev_comp[k] = delta(r1['por_severidade'].get(k, 0), r2['por_severidade'].get(k, 0))

    all_rules = set(list(r1['por_regra'].keys()) + list(r2['por_regra'].keys()))
    regra_comp = {
        regra: delta(r1['por_regra'].get(regra, 0), r2['por_regra'].get(regra, 0))
        for regra in all_rules
    }

    for r in [r1, r2]:
        r['download_url'] = f'/download/relatorio/{r["id"]}'
        r.pop('caminho_arquivo', None)

    return jsonify({
        'relatorio_a': r1,
        'relatorio_b': r2,
        'comparacao': {
            'total_lancamentos': delta(r1['total_lancamentos'], r2['total_lancamentos']),
            'total_erros':       delta(r1['total_erros'],       r2['total_erros']),
            'por_severidade':    sev_comp,
            'por_regra':         regra_comp,
        },
    })


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
