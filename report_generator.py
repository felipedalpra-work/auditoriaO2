"""
O2 Inc — Gerador de Relatório PDF de Auditoria
Branding Oficial: Vivid Green #6CF269 + Intense Gray #494949
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether,
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from datetime import datetime

PAGE_W, PAGE_H = A4

# ── Paleta O2 Inc (Branding Oficial) ────────────────────────────────────────
COR_VERDE       = colors.HexColor('#6CF269')   # Vivid Green — cor primária
COR_CINZA       = colors.HexColor('#494949')   # Intense Gray — cor secundária
COR_CINZA_CLARO = colors.HexColor('#F4F4F4')   # Fundos neutros
COR_CINZA_MEDIO = colors.HexColor('#E0E0E0')   # Linhas e bordas
COR_TEXTO       = colors.HexColor('#222222')   # Texto principal
COR_BRANCO      = colors.white

# Severidade (cores semânticas — mantidas por legibilidade funcional)
COR_CRITICO = colors.HexColor('#C0392B')
COR_ALTO    = colors.HexColor('#E67E22')
COR_MEDIO   = colors.HexColor('#B8860B')
COR_BAIXO   = colors.HexColor('#27AE60')

SEV_CORES = {
    'CRÍTICO': COR_CRITICO,
    'ALTO':    COR_ALTO,
    'MÉDIO':   COR_MEDIO,
    'BAIXO':   COR_BAIXO,
}
SEV_BG = {
    'CRÍTICO': colors.HexColor('#FDF0EF'),
    'ALTO':    colors.HexColor('#FEF7EF'),
    'MÉDIO':   colors.HexColor('#FDFBEF'),
    'BAIXO':   colors.HexColor('#EFF8F2'),
}


# ── Explicações por regra ─────────────────────────────────────────────────────
EXPLICACOES = {
    'Data de Emissão > Data de Vencimento': {
        'o_que_e': (
            'A data de emissão do documento é posterior à data de vencimento — '
            'situação matematicamente impossível. O item aparece como emitido depois '
            'de já estar vencido.'
        ),
        'por_que_importa': (
            'Afeta o correto reconhecimento de competência na plataforma Oxy. '
            'Pode mascarar inadimplências, antecipações ou distorcer o fluxo de '
            'caixa projetado do período, comprometendo análises de prazo e aging.'
        ),
        'como_corrigir': (
            'Consultar o documento fiscal original e corrigir a data de emissão ou '
            'de vencimento diretamente no Omie antes da próxima importação noturna.'
        ),
    },
    'Sinal de Valor Incorreto': {
        'o_que_e': (
            'Uma Conta a Pagar foi registrada com valor positivo, ou uma Conta a '
            'Receber com valor negativo. O sinal indica a direção do fluxo financeiro '
            'e está invertido em relação ao tipo do lançamento.'
        ),
        'por_que_importa': (
            'Uma despesa com sinal positivo aparece como receita na Oxy — e vice-versa. '
            'O erro contamina o resultado de toda a categoria afetada, distorcendo '
            'completamente o DRE e qualquer análise de receita versus despesa.'
        ),
        'como_corrigir': (
            'Corrigir o sinal do valor no lançamento dentro do Omie. '
            'Contas a Pagar devem ter valores negativos; Contas a Receber, positivos.'
        ),
    },
    'Lançamento Duplicado': {
        'o_que_e': (
            'Dois ou mais lançamentos possuem exatamente o mesmo CNPJ/CPF, valor, '
            'data de emissão e número de parcela. São registros idênticos que '
            'representam o mesmo fato contábil duplicado no sistema.'
        ),
        'por_que_importa': (
            'Cada duplicata multiplica o valor real no DRE. Um fornecedor com dois '
            'lançamentos idênticos aparece com o dobro do custo real, distorcendo '
            'análises de custo por categoria, por fornecedor e por período.'
        ),
        'como_corrigir': (
            'Identificar qual dos lançamentos é o correto e excluir os demais no Omie. '
            'Verificar se o processo de importação ou lançamento manual foi executado '
            'mais de uma vez para o mesmo documento fiscal.'
        ),
    },
    'Recorrência Quebrada': {
        'o_que_e': (
            'Um fornecedor que aparece regularmente em mais de 50% dos meses do período '
            'analisado está ausente em meses específicos, quebrando o padrão esperado '
            'de recorrência mensal.'
        ),
        'por_que_importa': (
            'Despesas fixas recorrentes (aluguel, assinaturas, serviços contínuos, folha) '
            'devem aparecer em todos os meses. A ausência pode indicar lançamento '
            'em mês errado ou uma despesa real que não foi registrada, criando lacunas '
            'no DRE e distorcendo análises mensais comparativas.'
        ),
        'como_corrigir': (
            'Verificar se o pagamento ocorreu no mês ausente. Se sim, localizar onde '
            'foi lançado e corrigir a data ou categoria. Se não ocorreu, avaliar se '
            'há justificativa válida e documentar.'
        ),
    },
    'Categoria Inconsistente por Fornecedor': {
        'o_que_e': (
            'O mesmo fornecedor foi classificado em categorias diferentes ao longo do '
            'período. O custo do serviço ou produto é o mesmo, mas o critério de '
            'categorização variou entre os lançamentos.'
        ),
        'por_que_importa': (
            'A Oxy agrupa despesas por categoria para construir o DRE. Categorização '
            'inconsistente distorce análises de tendência, invalida comparações entre '
            'meses e gera totais incorretos por categoria — comprometendo toda a '
            'análise gerencial baseada nesses números.'
        ),
        'como_corrigir': (
            'Definir a categoria correta para o fornecedor conforme o plano de contas '
            'da O2 e padronizar todos os lançamentos históricos. Considerar configurar '
            'regra de categorização automática no Omie para este fornecedor.'
        ),
    },
    'Valor Atípico (Outlier)': {
        'o_que_e': (
            'O valor deste lançamento desvia mais de 3 desvios-padrão da média histórica '
            'da sua categoria (z-score > 3), configurando um outlier estatístico '
            'altamente significativo em relação ao comportamento normal da categoria.'
        ),
        'por_que_importa': (
            'Pode indicar: erro de digitação (ex: R$ 1.500 lançado como R$ 15.000), '
            'duplicação parcial de valor, nota de outro período registrada no mês errado, '
            'ou um gasto extraordinário real que merece atenção gerencial. Em qualquer '
            'cenário, distorce a média e as análises de tendência da categoria.'
        ),
        'como_corrigir': (
            'Confirmar o valor no documento fiscal original. Se for erro de digitação, '
            'corrigir no Omie. Se for gasto extraordinário legítimo, documentar a '
            'justificativa na observação do lançamento para contexto nas análises futuras.'
        ),
    },
    'Parcelas Incompletas': {
        'o_que_e': (
            'Uma série de parcelas numeradas (ex: 1/12 a 12/12) apresenta lacunas — '
            'algumas parcelas esperadas não foram encontradas no período analisado, '
            'indicando descontinuidade na sequência.'
        ),
        'por_que_importa': (
            'Indica parcelas registradas em meses incorretos ou simplesmente não '
            'lançadas. Compromete projeções de fluxo de caixa, o custo total do '
            'contrato no DRE, a comparação entre períodos e o controle de '
            'obrigações financeiras futuras.'
        ),
        'como_corrigir': (
            'Localizar as parcelas faltantes no Omie. Se lançadas em datas erradas, '
            'corrigir o período de competência. Se não existirem, registrá-las no '
            'mês correto com a data de competência adequada.'
        ),
    },
    'Competência x Registro Distantes': {
        'o_que_e': (
            'A data de emissão do documento e a data de registro no Omie distam mais '
            'de 60 dias — o documento é antigo mas foi registrado muito depois no '
            'sistema (ou vice-versa).'
        ),
        'por_que_importa': (
            'Na plataforma Oxy, a data de registro é utilizada como data de '
            'competência. Uma diferença acima de 60 dias significa que a '
            'despesa ou receita está sendo reconhecida no mês errado do DRE, '
            'distorcendo análises comparativas e o resultado de cada período.'
        ),
        'como_corrigir': (
            'Corrigir a data de registro no Omie para refletir o mês em que o '
            'serviço foi prestado ou o bem foi entregue — independentemente de '
            'quando o pagamento foi processado ou a nota emitida.'
        ),
    },
    'Data de Registro Replicada (Bug Omie)': {
        'o_que_e': (
            'Todas as parcelas desta série têm exatamente a mesma data de registro — '
            'comportamento causado por um bug conhecido do Omie em notas de material '
            'parceladas, onde a data da primeira parcela é replicada automaticamente '
            'para todas as demais sem consentimento do usuário.'
        ),
        'por_que_importa': (
            'Como a Oxy usa a data de registro como competência, todas as parcelas '
            'são reconhecidas no mesmo mês, criando picos artificiais de despesa '
            'em um único período e ausências nos demais meses do contrato. '
            'Distorce completamente a curva de custos no DRE.'
        ),
        'como_corrigir': (
            'Corrigir manualmente a data de registro de cada parcela no Omie '
            '(o sistema não permite alteração em massa para notas de material). '
            'Um chamado técnico já foi aberto na Omie solicitando correção '
            'deste comportamento via campo específico de competência.'
        ),
    },
    'Juros não Separados da Amortização': {
        'o_que_e': (
            'Pagamento de empréstimo, financiamento ou consórcio lançado integralmente '
            'em uma única categoria, sem separar a parcela de juros (despesa financeira) '
            'da amortização do principal (saída patrimonial que não é despesa).'
        ),
        'por_que_importa': (
            'Juros impactam o resultado financeiro; amortização do principal não. '
            'Lançar tudo como despesa faz a empresa parecer mais deficitária do que é, '
            'distorcendo EBITDA, resultado financeiro e qualquer análise de '
            'endividamento e capacidade de pagamento no DRE.'
        ),
        'como_corrigir': (
            'Separar em dois lançamentos distintos no Omie: '
            '(1) valor dos juros → categoria de Despesas Financeiras; '
            '(2) valor da amortização → categoria de Empréstimos/Financiamentos. '
            'Os valores de cada componente constam no extrato do contrato '
            'com a instituição financeira.'
        ),
    },
}


# ── Utilitários ───────────────────────────────────────────────────────────────

def formatar_valor(v):
    try:
        f = float(v)
        return f'R$ {f:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return str(v)


def _hex(cor):
    """Retorna string hex sem # para uso em tags ReportLab."""
    h = cor.hexval()
    return h[2:] if h.startswith('0x') else h


def _sev_badge(sev):
    """Retorna HTML de badge de severidade colorido."""
    cor = SEV_CORES.get(sev, COR_CINZA)
    return f'<font color="#{_hex(cor)}"><b>{sev}</b></font>'


# ── Decoradores de página ─────────────────────────────────────────────────────

def _decorar_capa(canvas, doc):
    """Página de capa: fundo cinza escuro + faixas verdes."""
    canvas.saveState()
    canvas.setFillColor(COR_CINZA)
    canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    canvas.setFillColor(COR_VERDE)
    canvas.rect(0, PAGE_H - 8 * mm, PAGE_W, 8 * mm, fill=1, stroke=0)
    canvas.rect(0, 0, PAGE_W, 8 * mm, fill=1, stroke=0)
    canvas.restoreState()


def _decorar_interna(canvas, doc):
    """Páginas internas: faixa verde no topo + rodapé."""
    canvas.saveState()
    # Faixa verde superior
    canvas.setFillColor(COR_VERDE)
    canvas.rect(0, PAGE_H - 5 * mm, PAGE_W, 5 * mm, fill=1, stroke=0)
    # Linha separadora do rodapé
    canvas.setStrokeColor(COR_VERDE)
    canvas.setLineWidth(0.8)
    canvas.line(15 * mm, 14 * mm, PAGE_W - 15 * mm, 14 * mm)
    # Texto do rodapé
    canvas.setFillColor(COR_CINZA)
    canvas.setFont('Helvetica', 7)
    canvas.drawString(15 * mm, 9 * mm, 'O2 Inc — Assessoria Financeira  |  Documento Confidencial')
    canvas.drawRightString(PAGE_W - 15 * mm, 9 * mm, f'Página {doc.page}')
    canvas.restoreState()


# ── Função principal ──────────────────────────────────────────────────────────

def gerar_pdf(resultado, output_path, nome_arquivo_origem=''):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=15 * mm,
        leftMargin=15 * mm,
        topMargin=18 * mm,
        bottomMargin=22 * mm,
    )

    # ── Estilos ──────────────────────────────────────────────────────────────
    s = getSampleStyleSheet()

    # Capa
    st_capa_logo = ParagraphStyle('CapaLogo',
        fontSize=52, fontName='Helvetica-Bold',
        textColor=COR_VERDE, alignment=TA_LEFT, spaceAfter=2)

    st_capa_tagline = ParagraphStyle('CapaTagline',
        fontSize=13, fontName='Helvetica',
        textColor=COR_VERDE, alignment=TA_LEFT, spaceAfter=0)

    st_capa_titulo = ParagraphStyle('CapaTitulo',
        fontSize=24, fontName='Helvetica-Bold',
        textColor=COR_BRANCO, alignment=TA_LEFT, spaceAfter=4)

    st_capa_sub = ParagraphStyle('CapaSub',
        fontSize=13, fontName='Helvetica',
        textColor=colors.HexColor('#AAAAAA'), alignment=TA_LEFT, spaceAfter=3)

    st_capa_meta = ParagraphStyle('CapaMeta',
        fontSize=9, fontName='Helvetica',
        textColor=colors.HexColor('#888888'), alignment=TA_LEFT, spaceAfter=2)

    # Interno
    st_secao = ParagraphStyle('Secao',
        fontSize=12, fontName='Helvetica-Bold',
        textColor=COR_BRANCO, alignment=TA_LEFT,
        leftIndent=4, spaceAfter=0, spaceBefore=0)

    st_normal = ParagraphStyle('Normal2',
        fontSize=8, fontName='Helvetica',
        textColor=COR_TEXTO, spaceAfter=2, leading=11)

    st_cell = ParagraphStyle('Cell',
        fontSize=7, fontName='Helvetica',
        textColor=COR_TEXTO, leading=9)

    st_cell_bold = ParagraphStyle('CellBold',
        fontSize=7, fontName='Helvetica-Bold',
        textColor=COR_BRANCO, leading=9)

    st_expl_titulo = ParagraphStyle('ExplTitulo',
        fontSize=7, fontName='Helvetica-Bold',
        textColor=COR_CINZA, spaceAfter=3, leading=9)

    st_expl_texto = ParagraphStyle('ExplTexto',
        fontSize=7.5, fontName='Helvetica',
        textColor=COR_TEXTO, leading=10, spaceAfter=0)

    st_evidencia = ParagraphStyle('Evidencia',
        fontSize=7.5, fontName='Helvetica',
        textColor=COR_TEXTO, leading=10)

    story = []

    # ════════════════════════════════════════════════════════════════════
    # CAPA
    # ════════════════════════════════════════════════════════════════════
    story.append(Spacer(1, 28 * mm))
    story.append(Paragraph('O2INC', st_capa_logo))
    story.append(Paragraph('gestão fluída.', st_capa_tagline))
    story.append(Spacer(1, 18 * mm))
    story.append(HRFlowable(width='100%', thickness=1.5, color=COR_VERDE, spaceAfter=14))
    story.append(Paragraph('RELATÓRIO DE AUDITORIA', st_capa_titulo))
    story.append(Paragraph('DE LANÇAMENTOS OMIE', st_capa_sub))
    story.append(Spacer(1, 10 * mm))

    if nome_arquivo_origem:
        story.append(Paragraph(f'Arquivo analisado: {nome_arquivo_origem}', st_capa_meta))
    story.append(Paragraph(f'Gerado em: {resultado["data_analise"]}', st_capa_meta))

    story.append(Spacer(1, 24 * mm))

    # Boxes de estatísticas na capa
    total = resultado['total_lancamentos']
    total_erros = resultado['total_erros']
    pct = (total_erros / total * 100) if total > 0 else 0
    sev = resultado.get('por_severidade', {})

    def _stat_box(label, valor, cor_valor=COR_BRANCO):
        return Table(
            [
                [Paragraph(f'<font color="#{_hex(cor_valor)}"><b>{valor}</b></font>',
                    ParagraphStyle('SV', fontSize=22, fontName='Helvetica-Bold',
                        textColor=cor_valor, alignment=TA_CENTER))],
                [Paragraph(label,
                    ParagraphStyle('SL', fontSize=7.5, fontName='Helvetica',
                        textColor=colors.HexColor('#AAAAAA'), alignment=TA_CENTER, leading=10))],
            ],
            colWidths=[38 * mm],
        )

    stats_capa = [
        ['', '', '', ''],
        [
            _stat_box('Total Analisados', f'{total:,}'.replace(',', '.'), COR_BRANCO),
            _stat_box('Inconsistências', f'{total_erros}  ({pct:.1f}%)', COR_VERDE),
            _stat_box('Críticos', str(sev.get('CRÍTICO', 0)), COR_CRITICO),
            _stat_box('Alta Severidade', str(sev.get('ALTO', 0)), COR_ALTO),
        ],
    ]
    stat_table_capa = Table(stats_capa, colWidths=[38 * mm] * 4, hAlign='LEFT')
    stat_table_capa.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#3A3A3A')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#555555')),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        # Faixa verde no topo do box
        ('LINEABOVE', (0, 0), (-1, 0), 3, COR_VERDE),
    ]))
    story.append(stat_table_capa)
    story.append(PageBreak())

    # ════════════════════════════════════════════════════════════════════
    # SUMÁRIO EXECUTIVO
    # ════════════════════════════════════════════════════════════════════

    def _header_secao(texto):
        """Barra de seção cinza escuro com texto branco + faixa verde à esquerda."""
        tbl = Table([[Paragraph(texto, st_secao)]], colWidths=[180 * mm])
        tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), COR_CINZA),
            ('TOPPADDING', (0, 0), (-1, -1), 7),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('LINEBEFORE', (0, 0), (0, 0), 4, COR_VERDE),
        ]))
        return tbl

    story.append(_header_secao('SUMÁRIO EXECUTIVO'))
    story.append(Spacer(1, 5 * mm))

    # Tabela de totais por severidade
    resumo_data = [
        [
            Paragraph('<b>Indicador</b>', st_cell_bold),
            Paragraph('<b>Qtd</b>', st_cell_bold),
            Paragraph('<b>% do Total</b>', st_cell_bold),
        ],
        ['Total de Lançamentos Analisados', str(total), '100%'],
        ['Total de Inconsistências Encontradas', str(total_erros), f'{pct:.1f}%'],
    ]
    sev_order = [('CRÍTICO', COR_CRITICO), ('ALTO', COR_ALTO),
                 ('MÉDIO', COR_MEDIO), ('BAIXO', COR_BAIXO)]
    sev_styles_extra = []
    for row_i, (sev_nome, cor_sev) in enumerate(sev_order, start=len(resumo_data)):
        qtd_sev = sev.get(sev_nome, 0)
        pct_sev = (qtd_sev / total_erros * 100) if total_erros > 0 else 0
        resumo_data.append([
            Paragraph(f'  └ {sev_nome}',
                ParagraphStyle('SI', fontSize=8, fontName='Helvetica',
                    textColor=cor_sev, leading=10)),
            Paragraph(f'<b>{qtd_sev}</b>',
                ParagraphStyle('SN', fontSize=8, fontName='Helvetica-Bold',
                    textColor=cor_sev, leading=10, alignment=TA_CENTER)),
            f'{pct_sev:.1f}%',
        ])
        sev_styles_extra.append(
            ('BACKGROUND', (0, row_i), (-1, row_i), SEV_BG.get(sev_nome, COR_BRANCO))
        )

    resumo_tbl = Table(resumo_data, colWidths=[110 * mm, 30 * mm, 40 * mm])
    resumo_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), COR_CINZA),
        ('TEXTCOLOR', (0, 0), (-1, 0), COR_BRANCO),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('ROWBACKGROUNDS', (0, 1), (-1, 2), [COR_BRANCO, COR_CINZA_CLARO]),
        ('GRID', (0, 0), (-1, -1), 0.3, COR_CINZA_MEDIO),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
    ] + sev_styles_extra))
    story.append(resumo_tbl)
    story.append(Spacer(1, 8 * mm))

    # Tabela por regra
    por_regra = resultado.get('por_regra', {})
    if por_regra:
        story.append(_header_secao('INCONSISTÊNCIAS POR TIPO DE REGRA'))
        story.append(Spacer(1, 4 * mm))

        regra_data = [[
            Paragraph('<b>Regra de Auditoria</b>', st_cell_bold),
            Paragraph('<b>Ocorrências</b>', st_cell_bold),
            Paragraph('<b>% das Inconsist.</b>', st_cell_bold),
        ]]
        for regra_nome, qtd in sorted(por_regra.items(), key=lambda x: -x[1]):
            pct_r = (qtd / total_erros * 100) if total_erros > 0 else 0
            regra_data.append([regra_nome, str(qtd), f'{pct_r:.1f}%'])

        regra_tbl = Table(regra_data, colWidths=[110 * mm, 35 * mm, 35 * mm])
        regra_tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), COR_CINZA),
            ('TEXTCOLOR', (0, 0), (-1, 0), COR_BRANCO),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [COR_BRANCO, COR_CINZA_CLARO]),
            ('GRID', (0, 0), (-1, -1), 0.3, COR_CINZA_MEDIO),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ]))
        story.append(regra_tbl)

    story.append(PageBreak())

    # ════════════════════════════════════════════════════════════════════
    # SEÇÕES DETALHADAS POR REGRA
    # ════════════════════════════════════════════════════════════════════
    erros = resultado.get('erros', [])
    if not erros:
        story.append(_header_secao('DETALHAMENTO DAS INCONSISTÊNCIAS'))
        story.append(Spacer(1, 6 * mm))
        story.append(Paragraph('Nenhuma inconsistência encontrada nos lançamentos analisados.', st_normal))
    else:
        # Agrupa erros por regra, mantendo ordem de severidade
        from collections import defaultdict
        erros_por_regra = defaultdict(list)
        for e in erros:
            erros_por_regra[e['regra']].append(e)

        # Ordena as regras pela quantidade de erros (maior primeiro)
        regras_ordenadas = sorted(erros_por_regra.keys(), key=lambda r: -len(erros_por_regra[r]))

        for idx_regra, nome_regra in enumerate(regras_ordenadas):
            erros_regra = erros_por_regra[nome_regra]
            qtd = len(erros_regra)
            expl = EXPLICACOES.get(nome_regra, {})

            # ── Cabeçalho da seção de regra ──
            header_regra = Table(
                [[Paragraph(f'{nome_regra.upper()}  —  {qtd} ocorrência{"s" if qtd > 1 else ""}', st_secao)]],
                colWidths=[180 * mm]
            )
            header_regra.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), COR_CINZA),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
                ('LINEBEFORE', (0, 0), (0, 0), 4, COR_VERDE),
            ]))

            secao_elements = [header_regra, Spacer(1, 3 * mm)]

            # ── Caixa de explicação ──────────────────────────────────────
            if expl:
                expl_data = [
                    [
                        Paragraph('O QUE É', st_expl_titulo),
                        Paragraph('POR QUE IMPORTA', st_expl_titulo),
                        Paragraph('COMO CORRIGIR', st_expl_titulo),
                    ],
                    [
                        Paragraph(expl.get('o_que_e', ''), st_expl_texto),
                        Paragraph(expl.get('por_que_importa', ''), st_expl_texto),
                        Paragraph(expl.get('como_corrigir', ''), st_expl_texto),
                    ],
                ]
                expl_tbl = Table(expl_data, colWidths=[58 * mm, 62 * mm, 60 * mm])
                expl_tbl.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), COR_CINZA_CLARO),
                    ('LINEABOVE', (0, 0), (-1, 0), 2, COR_VERDE),
                    ('GRID', (0, 0), (-1, -1), 0.3, COR_CINZA_MEDIO),
                    ('TOPPADDING', (0, 0), (-1, -1), 6),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                    ('LEFTPADDING', (0, 0), (-1, -1), 7),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 7),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ]))
                secao_elements += [expl_tbl, Spacer(1, 4 * mm)]

            # ── Tabela de evidências ─────────────────────────────────────
            # Colunas: Sev | Fornecedor | CNPJ/CPF | Valor | Data Emissão | Categoria | Evidência
            col_widths = [13*mm, 35*mm, 23*mm, 19*mm, 16*mm, 24*mm, 50*mm]
            # total = 180mm

            header_ev = [
                Paragraph('<b>Sev.</b>', st_cell_bold),
                Paragraph('<b>Fornecedor / Cliente</b>', st_cell_bold),
                Paragraph('<b>CNPJ/CPF</b>', st_cell_bold),
                Paragraph('<b>Valor</b>', st_cell_bold),
                Paragraph('<b>Data</b>', st_cell_bold),
                Paragraph('<b>Categoria</b>', st_cell_bold),
                Paragraph('<b>Evidência</b>', st_cell_bold),
            ]

            ev_data = [header_ev]
            ev_styles = [
                ('BACKGROUND', (0, 0), (-1, 0), COR_CINZA),
                ('TEXTCOLOR', (0, 0), (-1, 0), COR_BRANCO),
                ('FONTSIZE', (0, 0), (-1, -1), 7),
                ('GRID', (0, 0), (-1, -1), 0.3, COR_CINZA_MEDIO),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]

            for row_i, erro in enumerate(erros_regra, start=1):
                sev_nome = erro.get('severidade', '')
                cor_sev = SEV_CORES.get(sev_nome, COR_CINZA)
                bg_row = SEV_BG.get(sev_nome, COR_BRANCO)

                fornecedor = str(erro.get('fornecedor', 'N/D'))[:40]
                cnpj = str(erro.get('cnpj', 'N/D'))
                valor = formatar_valor(erro.get('valor', 0))
                data_e = str(erro.get('data_emissao', 'N/D'))
                categoria = str(erro.get('categoria', 'N/D'))[:30]
                detalhe = str(erro.get('detalhe', ''))

                linha = [
                    Paragraph(
                        f'<font color="#{_hex(cor_sev)}"><b>{sev_nome}</b></font>',
                        st_cell
                    ),
                    Paragraph(fornecedor, st_cell),
                    Paragraph(cnpj, st_cell),
                    Paragraph(valor, st_cell),
                    Paragraph(data_e, st_cell),
                    Paragraph(categoria, st_cell),
                    Paragraph(detalhe, st_evidencia),
                ]
                ev_data.append(linha)

                ev_styles.append(('BACKGROUND', (0, row_i), (-1, row_i), bg_row))
                # Borda colorida de severidade na coluna de sev.
                ev_styles.append(('LINEBEFORE', (0, row_i), (0, row_i), 3, cor_sev))

            ev_tbl = Table(ev_data, colWidths=col_widths, repeatRows=1)
            ev_tbl.setStyle(TableStyle(ev_styles))

            secao_elements += [ev_tbl]

            # PageBreak entre seções (exceto a última)
            if idx_regra < len(regras_ordenadas) - 1:
                secao_elements.append(PageBreak())
            else:
                secao_elements.append(Spacer(1, 8 * mm))

            story.extend(secao_elements)

    # ── Rodapé final ─────────────────────────────────────────────────────────
    story.append(HRFlowable(width='100%', thickness=1, color=COR_VERDE, spaceAfter=6))
    story.append(Paragraph(
        f'O2 Inc — Assessoria Financeira  |  Relatório gerado em {resultado["data_analise"]}  |  Confidencial',
        ParagraphStyle('Rodape', fontSize=7, fontName='Helvetica',
            textColor=COR_CINZA, alignment=TA_CENTER)
    ))

    doc.build(story, onFirstPage=_decorar_capa, onLaterPages=_decorar_interna)
    return output_path
