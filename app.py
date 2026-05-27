"""
O2 Inc - Web App de Auditoria de Lançamentos Omie
"""

import os
import uuid
from dotenv import load_dotenv
load_dotenv()
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

from audit_engine import executar_auditoria
from excel_generator import gerar_excel

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

UPLOAD_FOLDER = '/tmp/o2_uploads'
REPORTS_DIR = '/tmp/o2_reports'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

EXTENSOES_PERMITIDAS = {'xlsx', 'xls'}


def permitido(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in EXTENSOES_PERMITIDAS


@app.errorhandler(RequestEntityTooLarge)
def handle_large_file(_err):
    return jsonify({'erro': 'Arquivo muito grande. Reduza o tamanho e tente novamente.'}), 413


# ── Páginas ──────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/historico')
def historico():
    return render_template('index.html')


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

    return jsonify({
        'sucesso':           True,
        'total_lancamentos': resultado['total_lancamentos'],
        'total_erros':       resultado['total_erros'],
        'por_severidade':    resultado['por_severidade'],
        'por_regra':         resultado['por_regra'],
        'download_url':      f'/download/arquivo/{nome_excel}',
        # Amostra para depuração rápida no frontend/API sem inflar demais o payload.
        'erros_debug':       resultado.get('erros', [])[:100],
    })


# ── Download ──────────────────────────────────────────────────────────────────

@app.route('/download/arquivo/<nome_arquivo>')
def download_arquivo(nome_arquivo):
    nome_safe = secure_filename(nome_arquivo)
    caminho = os.path.join(REPORTS_DIR, nome_safe)
    if not os.path.exists(caminho):
        return 'Arquivo não encontrado no servidor.', 404
    return send_file(
        caminho,
        as_attachment=True,
        download_name=nome_safe,
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
