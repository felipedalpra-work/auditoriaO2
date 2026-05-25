"""
O2 Inc - Engine de Auditoria de Lançamentos Omie
Detecta inconsistências e erros em relatórios de movimentação financeira.
"""

import pandas as pd
import numpy as np
from datetime import datetime
from collections import defaultdict
import unicodedata


def limpar_valor(v):
    if pd.isna(v) or v in ('N/D', 'Não informado', '', None):
        return None
    return str(v).strip()


def _normalizar_coluna(nome):
    txt = str(nome).strip().lower()
    txt = unicodedata.normalize('NFKD', txt).encode('ascii', 'ignore').decode('ascii')
    return ' '.join(txt.split())


def _aplicar_alias_colunas(df, aliases):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    colunas_norm = {_normalizar_coluna(c): c for c in df.columns}
    rename_map = {}
    for canonica, possiveis in aliases.items():
        if canonica in df.columns:
            continue
        for alias in possiveis:
            original = colunas_norm.get(_normalizar_coluna(alias))
            if original:
                rename_map[original] = canonica
                break
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def _linha_arquivo_row(row):
    v = row.get('__linha_arquivo')
    if pd.isna(v):
        return 'N/D'
    try:
        return int(v)
    except (TypeError, ValueError):
        return 'N/D'


def _linha_arquivo_grupo(df_grupo):
    if '__linha_arquivo' not in df_grupo.columns or df_grupo.empty:
        return 'N/D'
    try:
        return int(pd.to_numeric(df_grupo['__linha_arquivo'], errors='coerce').min())
    except (TypeError, ValueError):
        return 'N/D'


def _aba_arquivo_row(row):
    v = row.get('__aba_arquivo_origem')
    if pd.isna(v):
        return 'N/D'
    return str(v)


def _aba_arquivo_grupo(df_grupo):
    if '__aba_arquivo_origem' not in df_grupo.columns or df_grupo.empty:
        return 'N/D'
    v = df_grupo['__aba_arquivo_origem'].iloc[0]
    if pd.isna(v):
        return 'N/D'
    return str(v)


def carregar_df(filepath):
    aliases = {
        'Data de Emissão (completa)': [
            'data de emissao (completa)',
            'data de emissao completa',
            'data emissao (completa)',
            'data emissao completa',
            'data de emissao',
            'data emissao',
        ],
        'Tipo': [
            'tipo',
        ],
    }

    obrigatorias = ['Data de Emissão (completa)', 'Tipo']
    melhor_df = None
    melhor_faltantes = obrigatorias[:]

    xls = pd.ExcelFile(filepath)
    aba_origem = xls.sheet_names[0] if xls.sheet_names else 'Planilha1'

    # Omie pode exportar com variações de linhas acima do cabeçalho.
    for skip in (0, 1, 2, 3):
        cand = pd.read_excel(xls, sheet_name=0, skiprows=skip, header=0)
        cand['__linha_arquivo'] = cand.index + skip + 2
        cand['__aba_arquivo_origem'] = aba_origem
        cand = _aplicar_alias_colunas(cand, aliases)
        faltantes = [c for c in obrigatorias if c not in cand.columns]
        if not faltantes:
            melhor_df = cand
            break
        if len(faltantes) < len(melhor_faltantes):
            melhor_faltantes = faltantes
            melhor_df = cand

    df = melhor_df
    faltantes = [c for c in obrigatorias if c not in df.columns]
    if faltantes:
        raise ValueError(
            f"Colunas obrigatórias ausentes: {', '.join(faltantes)}. "
            "Verifique se o arquivo é o relatório padrão do Omie."
        )

    # Evita KeyError em cenários de planilha malformada e mantém limpeza básica.
    subset_dropna = [c for c in obrigatorias if c in df.columns]
    if subset_dropna:
        df = df.dropna(subset=subset_dropna, how='all')
    df = df[df['Tipo'].notna()]
    return df


# ── REGRAS CRÍTICAS ─────────────────────────────────────────────────────────

def regra_emissao_maior_vencimento(df):
    """Regra 1: Emissão mais de 60 dias após o vencimento."""
    erros = []
    col_e = 'Data de Emissão (completa)'
    col_v = 'Data de Vencimento (completa)'
    subset = df[df[col_e].notna() & df[col_v].notna()].copy()
    subset[col_e] = pd.to_datetime(subset[col_e], errors='coerce')
    subset[col_v] = pd.to_datetime(subset[col_v], errors='coerce')
    subset['_diff_dias'] = (subset[col_e] - subset[col_v]).dt.days
    problemas = subset[subset['_diff_dias'] > 60]
    for _, row in problemas.iterrows():
        dias = int(row['_diff_dias']) if pd.notna(row['_diff_dias']) else 0
        erros.append({
            'regra': 'Data de Emissão > Data de Vencimento (+60 dias)',
            'severidade': 'CRÍTICO',
            'linha_arquivo': _linha_arquivo_row(row),
            'aba_arquivo_origem': _aba_arquivo_row(row),
            'fornecedor': limpar_valor(row.get('Cliente ou Fornecedor (Razão Social)')) or 'N/D',
            'cnpj': limpar_valor(row.get('CNPJ/CPF')) or 'N/D',
            'valor': row.get('Valor da Conta', 0),
            'data_emissao': pd.to_datetime(row[col_e]).strftime('%d/%m/%Y') if pd.notna(row[col_e]) else 'N/D',
            'data_vencimento': pd.to_datetime(row[col_v]).strftime('%d/%m/%Y') if pd.notna(row[col_v]) else 'N/D',
            'categoria': limpar_valor(row.get('Categoria')) or 'N/D',
            'observacao': str(row.get('Observação da Conta', '')),
            'detalhe': f"Emissão {dias} dias após o vencimento; conferir competência e datas no lançamento.",
        })
    return erros


def regra_vencimento_igual_emissao(df):
    """Regra complementar: Data de vencimento igual à data de emissão."""
    erros = []
    col_e = 'Data de Emissão (completa)'
    col_v = 'Data de Vencimento (completa)'
    subset = df[df[col_e].notna() & df[col_v].notna()].copy()
    subset[col_e] = pd.to_datetime(subset[col_e], errors='coerce')
    subset[col_v] = pd.to_datetime(subset[col_v], errors='coerce')
    problemas = subset[subset[col_e] == subset[col_v]]
    for _, row in problemas.iterrows():
        data_ref = pd.to_datetime(row[col_e]).strftime('%d/%m/%Y') if pd.notna(row[col_e]) else 'N/D'
        erros.append({
            'regra': 'Data de Vencimento = Data de Emissão',
            'severidade': 'MÉDIO',
            'linha_arquivo': _linha_arquivo_row(row),
            'aba_arquivo_origem': _aba_arquivo_row(row),
            'fornecedor': limpar_valor(row.get('Cliente ou Fornecedor (Razão Social)')) or 'N/D',
            'cnpj': limpar_valor(row.get('CNPJ/CPF')) or 'N/D',
            'valor': row.get('Valor da Conta', 0),
            'data_emissao': data_ref,
            'data_vencimento': data_ref,
            'categoria': limpar_valor(row.get('Categoria')) or 'N/D',
            'observacao': str(row.get('Observação da Conta', '')),
            'detalhe': 'Vencimento igual à emissão; validar se o prazo está correto para a competência.',
        })
    return erros


def regra_vencimento_maior_emissao(df):
    """Regra complementar: Vencimento mais de 60 dias após emissão."""
    erros = []
    col_e = 'Data de Emissão (completa)'
    col_v = 'Data de Vencimento (completa)'
    subset = df[df[col_e].notna() & df[col_v].notna()].copy()
    subset[col_e] = pd.to_datetime(subset[col_e], errors='coerce')
    subset[col_v] = pd.to_datetime(subset[col_v], errors='coerce')
    subset['_diff_dias'] = (subset[col_v] - subset[col_e]).dt.days
    problemas = subset[subset['_diff_dias'] > 60]
    for _, row in problemas.iterrows():
        dias = int(row['_diff_dias']) if pd.notna(row['_diff_dias']) else 0
        erros.append({
            'regra': 'Data de Vencimento > Data de Emissão (+60 dias)',
            'severidade': 'BAIXO',
            'linha_arquivo': _linha_arquivo_row(row),
            'aba_arquivo_origem': _aba_arquivo_row(row),
            'fornecedor': limpar_valor(row.get('Cliente ou Fornecedor (Razão Social)')) or 'N/D',
            'cnpj': limpar_valor(row.get('CNPJ/CPF')) or 'N/D',
            'valor': row.get('Valor da Conta', 0),
            'data_emissao': pd.to_datetime(row[col_e]).strftime('%d/%m/%Y') if pd.notna(row[col_e]) else 'N/D',
            'data_vencimento': pd.to_datetime(row[col_v]).strftime('%d/%m/%Y') if pd.notna(row[col_v]) else 'N/D',
            'categoria': limpar_valor(row.get('Categoria')) or 'N/D',
            'observacao': str(row.get('Observação da Conta', '')),
            'detalhe': f'Vencimento {dias} dias após a emissão; conferir se prazo está compatível com a política financeira.',
        })
    return erros


def regra_sinal_errado(df):
    """Regra 2: Contas a Pagar com valor positivo ou Contas a Receber com valor negativo."""
    erros = []
    for _, row in df.iterrows():
        tipo = str(row.get('Tipo', '')).strip()
        try:
            valor = float(row.get('Valor da Conta', 0))
        except:
            continue
        problema = None
        if 'Pagar' in tipo and valor > 0:
            problema = f"Conta a Pagar com valor positivo: R$ {valor:,.2f}"
        elif 'Receber' in tipo and valor < 0:
            problema = f"Conta a Receber com valor negativo: R$ {valor:,.2f}"
        if problema:
            data_e = row.get('Data de Emissão (completa)')
            erros.append({
                'regra': 'Sinal de Valor Incorreto',
                'severidade': 'CRÍTICO',
                'linha_arquivo': _linha_arquivo_row(row),
                'aba_arquivo_origem': _aba_arquivo_row(row),
                'fornecedor': limpar_valor(row.get('Cliente ou Fornecedor (Razão Social)')) or 'N/D',
                'cnpj': limpar_valor(row.get('CNPJ/CPF')) or 'N/D',
                'valor': valor,
                'data_emissao': pd.to_datetime(data_e).strftime('%d/%m/%Y') if pd.notna(data_e) else 'N/D',
                'data_vencimento': 'N/D',
                'categoria': limpar_valor(row.get('Categoria')) or 'N/D',
                'observacao': str(row.get('Observação da Conta', '')),
                'detalhe': problema,
            })
    return erros


def regra_duplicatas(df):
    """Regra 3: Lançamentos duplicados (mesmo CNPJ + valor + data emissão + parcela).

    A parcela é incluída na chave para evitar falsos positivos: parcelas legítimas de
    um mesmo parcelamento (1/12, 2/12...) não devem ser sinalizadas como duplicatas.
    Apenas entradas com exatamente a mesma parcela são consideradas duplicatas reais.
    """
    erros = []
    col_e = 'Data de Emissão (completa)'
    df2 = df.copy()
    df2[col_e] = pd.to_datetime(df2[col_e], errors='coerce')

    # Inclui o número de parcela na chave para não confundir parcelas distintas com duplicatas
    if 'Parcela' in df2.columns:
        df2['_parcela_norm'] = df2['Parcela'].astype(str).str.strip().replace({'nan': '', 'N/D': '', 'None': ''})
    else:
        df2['_parcela_norm'] = ''

    df2['_key'] = (
        df2['CNPJ/CPF'].astype(str) + '||' +
        df2['Valor da Conta'].astype(str) + '||' +
        df2[col_e].dt.strftime('%Y-%m-%d') + '||' +
        df2['_parcela_norm']
    )
    counts = df2['_key'].value_counts()
    duplicadas = counts[counts > 1].index
    vistos = set()
    for _, row in df2[df2['_key'].isin(duplicadas)].iterrows():
        key = row['_key']
        if key in vistos:
            continue
        vistos.add(key)
        qtd = counts[key]
        data_e = row[col_e]
        parcela_info = f", parcela {row['_parcela_norm']}" if row['_parcela_norm'] else ""
        erros.append({
            'regra': 'Lançamento Duplicado',
            'severidade': 'ALTO',
            'linha_arquivo': _linha_arquivo_row(row),
            'aba_arquivo_origem': _aba_arquivo_row(row),
            'fornecedor': limpar_valor(row.get('Cliente ou Fornecedor (Razão Social)')) or 'N/D',
            'cnpj': limpar_valor(row.get('CNPJ/CPF')) or 'N/D',
            'valor': row.get('Valor da Conta', 0),
            'data_emissao': data_e.strftime('%d/%m/%Y') if pd.notna(data_e) else 'N/D',
            'data_vencimento': 'N/D',
            'categoria': limpar_valor(row.get('Categoria')) or 'N/D',
            'observacao': str(row.get('Observação da Conta', '')),
            'detalhe': f"Aparece {qtd}x com mesmo CNPJ, valor e data{parcela_info}",
        })
    return erros


def regra_recorrencia_quebrada(df):
    """Regra 4: Fornecedor recorrente com mês ausente no período."""
    erros = []
    col_e = 'Data de Emissão (completa)'
    df2 = df.copy()
    df2[col_e] = pd.to_datetime(df2[col_e], errors='coerce')
    df2['_mes'] = df2[col_e].dt.to_period('M')
    df2['_fornecedor'] = df2['Cliente ou Fornecedor (Razão Social)'].astype(str)

    meses_no_periodo = set(df2['_mes'].dropna().unique())
    total_meses = len(meses_no_periodo)
    threshold = max(2, int(total_meses * 0.5))

    fornecedor_meses = df2.groupby('_fornecedor')['_mes'].apply(set)
    for fornecedor, meses_presentes in fornecedor_meses.items():
        if fornecedor in ('N/D', 'nan', ''):
            continue
        if len(meses_presentes) >= threshold:
            meses_ausentes = meses_no_periodo - meses_presentes
            if meses_ausentes:
                meses_str = ', '.join(sorted(str(m) for m in meses_ausentes))
                valor_medio = df2[df2['_fornecedor'] == fornecedor]['Valor da Conta'].mean()
                grupo_f = df2[df2['_fornecedor'] == fornecedor]
                cnpj = grupo_f['CNPJ/CPF'].iloc[0]
                categoria = grupo_f['Categoria'].iloc[0]
                erros.append({
                    'regra': 'Recorrência Quebrada',
                    'severidade': 'MÉDIO',
                    'linha_arquivo': _linha_arquivo_grupo(grupo_f),
                    'aba_arquivo_origem': _aba_arquivo_grupo(grupo_f),
                    'fornecedor': fornecedor,
                    'cnpj': limpar_valor(cnpj) or 'N/D',
                    'valor': valor_medio,
                    'data_emissao': 'Vários',
                    'data_vencimento': 'N/D',
                    'categoria': limpar_valor(categoria) or 'N/D',
                    'observacao': '',
                    'detalhe': f"Recorrente em {len(meses_presentes)} meses, ausente em: {meses_str}",
                })
    return erros


# ── NOVAS REGRAS ─────────────────────────────────────────────────────────────

def regra_categoria_inconsistente(df):
    """Regra 5: Mesmo fornecedor categorizado de formas diferentes entre meses."""
    erros = []
    col_e = 'Data de Emissão (completa)'
    df2 = df.copy()
    df2[col_e] = pd.to_datetime(df2[col_e], errors='coerce')
    df2['_fornecedor'] = df2['Cliente ou Fornecedor (Razão Social)'].astype(str)
    df2['_cat'] = df2['Categoria'].astype(str)

    for fornecedor, grupo in df2.groupby('_fornecedor'):
        if fornecedor in ('N/D', 'nan', ''):
            continue
        categorias = set(grupo['_cat'].unique()) - {'N/D', 'nan', ''}
        if len(categorias) > 1:
            cats_str = ' / '.join(sorted(categorias))
            cnpj = grupo['CNPJ/CPF'].iloc[0]
            valor_total = grupo['Valor da Conta'].sum()
            erros.append({
                'regra': 'Categoria Inconsistente por Fornecedor',
                'severidade': 'ALTO',
                'linha_arquivo': _linha_arquivo_grupo(grupo),
                'aba_arquivo_origem': _aba_arquivo_grupo(grupo),
                'fornecedor': fornecedor,
                'cnpj': limpar_valor(cnpj) or 'N/D',
                'valor': valor_total,
                'data_emissao': 'Vários',
                'data_vencimento': 'N/D',
                'categoria': cats_str[:60],
                'observacao': '',
                'detalhe': f"Lançado em {len(categorias)} categorias diferentes: {cats_str[:80]}",
            })
    return erros


def regra_outlier_valor(df):
    """Regra 6: Valor muito acima da média histórica do próprio fornecedor (z-score > 3)."""
    erros = []
    col_e = 'Data de Emissão (completa)'
    df2 = df.copy()
    df2[col_e] = pd.to_datetime(df2[col_e], errors='coerce')
    df2['_valor_abs'] = pd.to_numeric(df2['Valor da Conta'], errors='coerce').abs()
    df2['_fornecedor'] = df2['Cliente ou Fornecedor (Razão Social)'].astype(str)

    # Analisa outliers por fornecedor (mínimo 5 lançamentos para ter base estatística)
    for fornecedor, grupo in df2.groupby('_fornecedor'):
        if fornecedor in ('N/D', 'nan', '') or len(grupo) < 5:
            continue
        media = grupo['_valor_abs'].mean()
        std = grupo['_valor_abs'].std()
        if std == 0:
            continue
        # z-score > 3 = outlier
        outliers = grupo[((grupo['_valor_abs'] - media) / std).abs() > 3]
        for _, row in outliers.iterrows():
            val = row['_valor_abs']
            data_e = row[col_e]
            erros.append({
                'regra': 'Valor Atípico (Outlier)',
                'severidade': 'ALTO',
                'linha_arquivo': _linha_arquivo_row(row),
                'aba_arquivo_origem': _aba_arquivo_row(row),
                'fornecedor': limpar_valor(row.get('Cliente ou Fornecedor (Razão Social)')) or 'N/D',
                'cnpj': limpar_valor(row.get('CNPJ/CPF')) or 'N/D',
                'valor': row.get('Valor da Conta', 0),
                'data_emissao': data_e.strftime('%d/%m/%Y') if pd.notna(data_e) else 'N/D',
                'data_vencimento': 'N/D',
                'categoria': limpar_valor(row.get('Categoria')) or 'N/D',
                'observacao': str(row.get('Observação da Conta', '')),
                'detalhe': f"Valor R$ {val:,.2f} é outlier — média do fornecedor R$ {media:,.2f} (±{std:,.2f})",
            })
    return erros


def regra_parcelas_incompletas(df):
    """Regra 7: Série de parcelas com lacunas (ex: tem 1/12 e 3/12 mas falta 2/12)."""
    erros = []
    col_parcela = 'Parcela'
    if col_parcela not in df.columns:
        return erros

    df2 = df.copy()
    df2['_fornecedor'] = df2['Cliente ou Fornecedor (Razão Social)'].astype(str)
    df2['_parcela_str'] = df2[col_parcela].astype(str).str.strip()

    # Extrai número_atual e total de "NNN/TTT"
    mask = df2['_parcela_str'].str.match(r'^\d+/\d+$')
    df2 = df2[mask].copy()
    if df2.empty:
        return erros

    df2['_num'] = df2['_parcela_str'].str.split('/').str[0].astype(int)
    df2['_tot'] = df2['_parcela_str'].str.split('/').str[1].astype(int)

    # Agrupa por fornecedor + total de parcelas
    grupos = df2.groupby(['_fornecedor', 'CNPJ/CPF', 'Categoria', '_tot'])
    for (fornecedor, cnpj, cat, total), grupo in grupos:
        if fornecedor in ('N/D', 'nan', '') or total < 2:
            continue
        parcelas_presentes = set(grupo['_num'].tolist())
        esperadas = set(range(1, total + 1))
        faltando = esperadas - parcelas_presentes
        if faltando:
            faltando_str = ', '.join(str(p) for p in sorted(faltando))
            valor_medio = grupo['Valor da Conta'].abs().mean()
            erros.append({
                'regra': 'Parcelas Incompletas',
                'severidade': 'MÉDIO',
                'linha_arquivo': _linha_arquivo_grupo(grupo),
                'aba_arquivo_origem': _aba_arquivo_grupo(grupo),
                'fornecedor': fornecedor,
                'cnpj': limpar_valor(cnpj) or 'N/D',
                'valor': valor_medio,
                'data_emissao': 'Vários',
                'data_vencimento': 'N/D',
                'categoria': limpar_valor(cat) or 'N/D',
                'observacao': '',
                'detalhe': f"Série de {total} parcelas — faltam as parcelas: {faltando_str}",
            })
    return erros


def regra_competencia_distante(df, usar_registro_como_competencia=False):
    """Regra 8: Data de emissão muito distante da data de registro (>60 dias)."""
    erros = []
    if usar_registro_como_competencia:
        # Para clientes cuja competência oficial é a data de registro,
        # essa regra não se aplica (evita falso positivo sistêmico).
        return erros

    col_e = 'Data de Emissão (completa)'
    col_r = 'Data de Registro (completa)'
    if col_r not in df.columns:
        return erros

    df2 = df.copy()
    df2[col_e] = pd.to_datetime(df2[col_e], errors='coerce')
    df2[col_r] = pd.to_datetime(df2[col_r], errors='coerce')
    mask = df2[col_e].notna() & df2[col_r].notna()
    df2 = df2[mask].copy()
    df2['_diff'] = (df2[col_r] - df2[col_e]).dt.days.abs()

    problemas = df2[df2['_diff'] > 60]
    for _, row in problemas.iterrows():
        dias = int(row['_diff'])
        erros.append({
            'regra': 'Competência x Registro Distantes',
            'severidade': 'MÉDIO',
            'linha_arquivo': _linha_arquivo_row(row),
            'aba_arquivo_origem': _aba_arquivo_row(row),
            'fornecedor': limpar_valor(row.get('Cliente ou Fornecedor (Razão Social)')) or 'N/D',
            'cnpj': limpar_valor(row.get('CNPJ/CPF')) or 'N/D',
            'valor': row.get('Valor da Conta', 0),
            'data_emissao': row[col_e].strftime('%d/%m/%Y'),
            'data_vencimento': row[col_r].strftime('%d/%m/%Y'),
            'categoria': limpar_valor(row.get('Categoria')) or 'N/D',
            'observacao': str(row.get('Observação da Conta', '')),
            'detalhe': f"Emissão e registro distam {dias} dias — pode indicar lançamento no mês errado",
        })
    return erros


def regra_parcelas_data_registro_identica(df):
    """Regra 9: Parcelas com data de registro idêntica — bug conhecido do Omie.

    Notas de material com parcelamento no Omie replicam automaticamente a data de
    registro da primeira parcela para todas as demais sem consentimento do usuário.
    Quando uma série completa de parcelas tem exatamente a mesma data de registro,
    é altamente provável que seja este bug, não um lançamento intencional.
    """
    erros = []
    col_parcela = 'Parcela'
    col_r = 'Data de Registro (completa)'
    if col_parcela not in df.columns or col_r not in df.columns:
        return erros

    df2 = df.copy()
    df2['_fornecedor'] = df2['Cliente ou Fornecedor (Razão Social)'].astype(str)
    df2['_parcela_str'] = df2[col_parcela].astype(str).str.strip()
    df2[col_r] = pd.to_datetime(df2[col_r], errors='coerce')

    # Apenas linhas com formato válido de parcela (N/T)
    mask = df2['_parcela_str'].str.match(r'^\d+/\d+$')
    df2 = df2[mask].copy()
    if df2.empty:
        return erros

    df2['_tot'] = df2['_parcela_str'].str.split('/').str[1].astype(int)

    # Agrupa por fornecedor + CNPJ + categoria + total de parcelas
    grupos = df2.groupby(['_fornecedor', 'CNPJ/CPF', 'Categoria', '_tot'])
    for (fornecedor, cnpj, cat, total), grupo in grupos:
        if fornecedor in ('N/D', 'nan', '') or total < 2:
            continue
        # Verifica se todas as parcelas têm a mesma data de registro
        datas = grupo[col_r].dropna().unique()
        if len(datas) == 1 and len(grupo) >= total:
            data_str = pd.to_datetime(datas[0]).strftime('%d/%m/%Y')
            valor_medio = grupo['Valor da Conta'].abs().mean()
            erros.append({
                'regra': 'Data de Registro Replicada (Bug Omie)',
                'severidade': 'MÉDIO',
                'linha_arquivo': _linha_arquivo_grupo(grupo),
                'aba_arquivo_origem': _aba_arquivo_grupo(grupo),
                'fornecedor': fornecedor,
                'cnpj': limpar_valor(cnpj) or 'N/D',
                'valor': valor_medio,
                'data_emissao': 'Vários',
                'data_vencimento': 'N/D',
                'categoria': limpar_valor(cat) or 'N/D',
                'observacao': '',
                'detalhe': (
                    f"Série de {total} parcelas com data de registro idêntica ({data_str}). "
                    f"Provável replicação automática do Omie em nota de material — "
                    f"a data de competência de cada parcela pode estar incorreta."
                ),
            })
    return erros


def regra_juros_sem_amortizacao(df):
    """Regra 10: Pagamento de empréstimo/financiamento sem separar juros da amortização.

    Quando um pagamento de empréstimo, financiamento ou consórcio é lançado em uma
    única categoria, os juros (despesa financeira) ficam misturados com a amortização
    do principal (não é despesa). Isso distorce o resultado financeiro no DRE.

    Como o Omie não tem tipo de gasto padrão para juros de empréstimo, esta regra
    usa palavras-chave nas categorias e nomes de fornecedores para identificar os casos.
    """
    erros = []

    # Palavras que indicam categoria de principal de empréstimo
    KEYWORDS_PRINCIPAL = [
        'empréstimo', 'emprestimo', 'financiamento', 'consórcio', 'consorcio',
        'amortização', 'amortizacao', 'leasing', 'arrendamento mercantil',
    ]
    # Palavras que indicam categoria de juros/encargos financeiros
    KEYWORDS_JUROS = [
        'juro', 'juros', 'encargo financeiro', 'encargos financeiros',
        'despesa financeira', 'despesas financeiras', 'mora', 'iof',
    ]

    col_e = 'Data de Emissão (completa)'
    df2 = df.copy()
    df2[col_e] = pd.to_datetime(df2[col_e], errors='coerce')
    df2['_mes'] = df2[col_e].dt.to_period('M')
    df2['_cat_lower'] = df2['Categoria'].astype(str).str.lower()
    df2['_forn'] = df2['Cliente ou Fornecedor (Razão Social)'].astype(str)

    def tem_keyword(s, keywords):
        s = str(s).lower()
        return any(kw in s for kw in keywords)

    # Filtra apenas entradas que indicam principal de empréstimo
    mask_principal = df2['_cat_lower'].apply(lambda s: tem_keyword(s, KEYWORDS_PRINCIPAL))
    suspeitos = df2[mask_principal].copy()
    if suspeitos.empty:
        return erros

    vistos = set()
    for (fornecedor, mes), grupo in suspeitos.groupby(['_forn', '_mes']):
        key = (str(fornecedor), str(mes))
        if key in vistos or fornecedor in ('', 'nan', 'N/D'):
            continue
        vistos.add(key)

        # Verifica se este fornecedor tem algum lançamento de juros no mesmo mês
        todos_do_mes = df2[(df2['_forn'] == fornecedor) & (df2['_mes'] == mes)]
        has_juros = todos_do_mes['_cat_lower'].apply(
            lambda s: tem_keyword(s, KEYWORDS_JUROS)
        ).any()

        if not has_juros:
            total_valor = grupo['Valor da Conta'].abs().sum()
            if total_valor < 500:  # Ignora valores muito baixos
                continue
            cnpj = grupo['CNPJ/CPF'].iloc[0]
            cat = grupo['Categoria'].iloc[0]
            num_lanc = len(grupo)
            erros.append({
                'regra': 'Juros não Separados da Amortização',
                'severidade': 'ALTO',
                'linha_arquivo': _linha_arquivo_grupo(grupo),
                'aba_arquivo_origem': _aba_arquivo_grupo(grupo),
                'fornecedor': fornecedor,
                'cnpj': limpar_valor(cnpj) or 'N/D',
                'valor': total_valor,
                'data_emissao': str(mes),
                'data_vencimento': 'N/D',
                'categoria': limpar_valor(cat) or 'N/D',
                'observacao': '',
                'detalhe': (
                    f"{num_lanc} lançamento(s) de '{cat}' em {mes} sem registro separado de juros. "
                    f"Juros devem ser lançados em categoria de despesa financeira e "
                    f"amortização do principal em categoria de investimento/passivo."
                ),
            })
    return erros


def regra_categoria_semantica(df):
    """Regra 11: Fornecedor categorizado em categoria semanticamente incorreta (via OpenAI).

    Extrai todos os pares únicos (fornecedor, categoria) e envia em uma única chamada
    ao GPT-4o-mini para identificar combinações que não fazem sentido semântico.
    Se a chave de API não estiver configurada, a regra é silenciosamente ignorada.
    """
    import os
    import json

    try:
        from openai import OpenAI
    except ImportError:
        return []

    api_key = os.getenv('OPENAI_API_KEY', '').strip()
    if not api_key or api_key.startswith('sk-...'):
        return []

    client = OpenAI(api_key=api_key)

    df2 = df.copy()
    df2['_fornecedor'] = df2['Cliente ou Fornecedor (Razão Social)'].astype(str).str.strip()
    df2['_cat']        = df2['Categoria'].astype(str).str.strip()
    col_e = 'Data de Emissão (completa)'

    # Pares únicos com filtro de valores inválidos
    invalidos = {'nan', 'N/D', 'None', '', 'Não informado'}
    pares_df = (
        df2[~df2['_fornecedor'].isin(invalidos) & ~df2['_cat'].isin(invalidos)]
        .groupby(['_fornecedor', '_cat'])
        .size()
        .reset_index(name='_qtd')
    )

    if pares_df.empty:
        return []

    CHUNK = 80  # pares por chamada para não estourar o contexto
    pares_errados = []

    for inicio in range(0, len(pares_df), CHUNK):
        chunk = pares_df.iloc[inicio: inicio + CHUNK]
        lista = [
            {'fornecedor': r['_fornecedor'], 'categoria': r['_cat']}
            for _, r in chunk.iterrows()
        ]

        prompt = (
            'Você é um auditor financeiro brasileiro. '
            'Analise os pares de fornecedor e categoria de lançamento contábil abaixo.\n'
            'Retorne APENAS os pares onde a categoria está claramente errada para aquele fornecedor.\n'
            'Seja conservador: sinalize somente quando tiver alta confiança. '
            'Ignore casos ambíguos.\n\n'
            f'Pares:\n{json.dumps(lista, ensure_ascii=False)}\n\n'
            'Responda em JSON: {"inconsistencias": [{"fornecedor": "...", "categoria": "...", "motivo": "..."}]}'
        )

        try:
            resp = client.chat.completions.create(
                model='gpt-4o-mini',
                messages=[{'role': 'user', 'content': prompt}],
                response_format={'type': 'json_object'},
                temperature=0,
            )
            resultado = json.loads(resp.choices[0].message.content)
            pares_errados.extend(resultado.get('inconsistencias', []))
        except Exception:
            continue

    erros = []
    for par in pares_errados:
        forn   = par.get('fornecedor', '')
        cat    = par.get('categoria', '')
        motivo = par.get('motivo', '')
        mask   = (df2['_fornecedor'] == forn) & (df2['_cat'] == cat)
        for _, row in df2[mask].iterrows():
            data_e = pd.to_datetime(row.get(col_e), errors='coerce')
            erros.append({
                'regra':          'Categoria Semanticamente Incorreta',
                'severidade':     'ALTO',
                'linha_arquivo':  _linha_arquivo_row(row),
                'aba_arquivo_origem': _aba_arquivo_row(row),
                'fornecedor':     limpar_valor(row.get('Cliente ou Fornecedor (Razão Social)')) or 'N/D',
                'cnpj':           limpar_valor(row.get('CNPJ/CPF')) or 'N/D',
                'valor':          row.get('Valor da Conta', 0),
                'data_emissao':   data_e.strftime('%d/%m/%Y') if pd.notna(data_e) else 'N/D',
                'data_vencimento':'N/D',
                'categoria':      cat,
                'observacao':     str(row.get('Observação da Conta', '')),
                'detalhe':        f'Categoria "{cat}" parece incorreta para este fornecedor — {motivo}',
            })
    return erros


# ── EXECUTOR ─────────────────────────────────────────────────────────────────

ORDEM_SEVERIDADE = {'CRÍTICO': 0, 'ALTO': 1, 'MÉDIO': 2, 'BAIXO': 3}

REGRAS = [
    regra_emissao_maior_vencimento,
    regra_vencimento_igual_emissao,
    regra_vencimento_maior_emissao,
    regra_sinal_errado,
    regra_duplicatas,
    regra_recorrencia_quebrada,
    regra_categoria_inconsistente,
    regra_outlier_valor,
    regra_parcelas_incompletas,
    regra_competencia_distante,
    regra_parcelas_data_registro_identica,
    regra_juros_sem_amortizacao,
    regra_categoria_semantica,  # requer OPENAI_API_KEY no .env
]


def executar_auditoria(filepath, usar_registro_como_competencia=False):
    df = carregar_df(filepath)
    total_lancamentos = len(df)
    todos_erros = []
    resumo_por_regra = {}

    for regra_fn in REGRAS:
        try:
            if regra_fn.__name__ == 'regra_competencia_distante':
                erros = regra_fn(df, usar_registro_como_competencia=usar_registro_como_competencia)
            else:
                erros = regra_fn(df)
            todos_erros.extend(erros)
            if erros:
                resumo_por_regra[erros[0]['regra']] = len(erros)
        except Exception as e:
            pass

    todos_erros.sort(key=lambda x: ORDEM_SEVERIDADE.get(x['severidade'], 99))

    por_severidade = defaultdict(int)
    for e in todos_erros:
        por_severidade[e['severidade']] += 1

    return {
        'total_lancamentos': total_lancamentos,
        'total_erros': len(todos_erros),
        'por_severidade': dict(por_severidade),
        'por_regra': resumo_por_regra,
        'erros': todos_erros,
        'data_analise': datetime.now().strftime('%d/%m/%Y %H:%M'),
    }
