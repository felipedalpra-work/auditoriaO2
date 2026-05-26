"""
O2 Inc — Persistência SQLite para histórico de auditorias por empresa.
"""
import sqlite3
import json
import os

DATA_DIR    = os.getenv('AUDITORIA_DATA_DIR', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data'))
DB_PATH     = os.path.join(DATA_DIR, 'auditoria.db')
REPORTS_DIR = os.path.join(DATA_DIR, 'reports')


def init_db():
    os.makedirs(REPORTS_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute('''
        CREATE TABLE IF NOT EXISTS relatorios (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa             TEXT    NOT NULL,
            data_analise        TEXT    NOT NULL,
            nome_arquivo_origem TEXT    NOT NULL,
            total_lancamentos   INTEGER DEFAULT 0,
            total_erros         INTEGER DEFAULT 0,
            por_severidade      TEXT    DEFAULT '{}',
            por_regra           TEXT    DEFAULT '{}',
            caminho_arquivo     TEXT    NOT NULL
        )
    ''')
    conn.commit()
    conn.close()


def salvar_relatorio(empresa, data_analise, nome_arquivo_origem,
                     total_lancamentos, total_erros,
                     por_severidade, por_regra, caminho_arquivo):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        '''INSERT INTO relatorios
           (empresa, data_analise, nome_arquivo_origem, total_lancamentos, total_erros,
            por_severidade, por_regra, caminho_arquivo)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
        (empresa, data_analise, nome_arquivo_origem,
         total_lancamentos, total_erros,
         json.dumps(por_severidade, ensure_ascii=False),
         json.dumps(por_regra,      ensure_ascii=False),
         caminho_arquivo)
    )
    conn.commit()
    rid = cur.lastrowid
    conn.close()
    return rid


def listar_empresas():
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        '''SELECT empresa, COUNT(*) AS qtd, MAX(data_analise) AS ultima
           FROM relatorios
           GROUP BY empresa
           ORDER BY ultima DESC'''
    ).fetchall()
    conn.close()
    return [
        {'empresa': r[0], 'total_relatorios': r[1], 'ultima_analise': r[2]}
        for r in rows
    ]


def listar_relatorios_empresa(empresa):
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        '''SELECT id, empresa, data_analise, nome_arquivo_origem,
                  total_lancamentos, total_erros, por_severidade, por_regra, caminho_arquivo
           FROM relatorios
           WHERE empresa = ?
           ORDER BY data_analise DESC''',
        (empresa,)
    ).fetchall()
    conn.close()
    return [_row_to_dict(r) for r in rows]


def buscar_relatorio(rid):
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        '''SELECT id, empresa, data_analise, nome_arquivo_origem,
                  total_lancamentos, total_erros, por_severidade, por_regra, caminho_arquivo
           FROM relatorios WHERE id = ?''',
        (rid,)
    ).fetchone()
    conn.close()
    return _row_to_dict(row) if row else None


def _row_to_dict(r):
    return {
        'id':                  r[0],
        'empresa':             r[1],
        'data_analise':        r[2],
        'nome_arquivo_origem': r[3],
        'total_lancamentos':   r[4],
        'total_erros':         r[5],
        'por_severidade':      json.loads(r[6]) if r[6] else {},
        'por_regra':           json.loads(r[7]) if r[7] else {},
        'caminho_arquivo':     r[8],
    }
