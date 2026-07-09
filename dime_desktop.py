#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dime_desktop.py
================

Programa DESKTOP (janela gráfica) que converte o arquivo-texto de layout
fixo da DIME/GIA-SC diretamente em PDF, no formato "DIME Detalhada",
conforme o layout descrito no Manual Consolidado da DIME (Anexo I -
Layout dos Registros).

Este programa reaproveita todo o motor de leitura/validação já testado em
"dime_txt_to_detalhada.py" (mesma pasta) e apenas adiciona:
  1) geração de PDF nativa (biblioteca reportlab, sem depender de
     LibreOffice/Word instalados na máquina);
  2) uma interface gráfica (tkinter, já incluído no Python padrão do
     Windows) para selecionar o arquivo de entrada e a pasta de destino
     com cliques de mouse.

INSTALAÇÃO (uma única vez)
---------------------------
    pip install reportlab

EXECUÇÃO
--------
    python dime_desktop.py

Também é possível gerar um executável único (.exe) para Windows, sem
precisar instalar Python na máquina de destino:

    pip install pyinstaller
    pyinstaller --onefile --windowed --name "Conversor DIME" dime_desktop.py

O executável final aparecerá em dist\\Conversor DIME.exe.
"""

import os
import sys
import threading
import traceback

# ---------------------------------------------------------------------------
# Motor de leitura/validação (compartilhado com a versão de linha de comando)
# ---------------------------------------------------------------------------
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dime_txt_to_detalhada as core  # noqa: E402


# ---------------------------------------------------------------------------
# Geração de PDF nativa (reportlab) - não depende de LibreOffice/Word
# ---------------------------------------------------------------------------

def _import_reportlab():
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.units import mm
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                         Paragraph, Spacer, PageBreak)
        return dict(colors=colors, A4=A4, landscape=landscape, mm=mm,
                    getSampleStyleSheet=getSampleStyleSheet,
                    SimpleDocTemplate=SimpleDocTemplate, Table=Table,
                    TableStyle=TableStyle, Paragraph=Paragraph,
                    Spacer=Spacer, PageBreak=PageBreak)
    except ImportError as exc:
        raise ImportError(
            "A biblioteca 'reportlab' não está instalada.\n"
            "Instale com:  pip install reportlab"
        ) from exc


def gerar_pdf(contador, decl, caminho_pdf, log=print):
    """Gera o relatório DIME Detalhada em PDF a partir de uma declaração já
    processada por dime_txt_to_detalhada.parse_file(). Reproduz o mais
    fielmente possível o layout do extrato oficial "DIME Detalhada" emitido
    pelo sistema S@T da SEF/SC (cabeçalho, caixa de índice "QUADROS", cada
    quadro como uma única caixa com título + link "Topo", cabeçalhos
    mesclados nos Quadros 01/02 etc.). Não depende de tkinter - pode ser
    usada em lote/linha de comando."""
    rl = _import_reportlab()
    colors, A4, mm = rl['colors'], rl['A4'], rl['mm']
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

    q00 = decl['quadro00']
    ie = core.fmt_ie(q00['inscricao'])
    nome = q00['nome']
    periodo = core.fmt_periodo(q00['periodo'])

    PAGE_W, _ = A4
    MARGEM = 11 * mm
    LARGURA = PAGE_W - 2 * MARGEM

    doc = rl['SimpleDocTemplate'](
        caminho_pdf, pagesize=A4,
        leftMargin=MARGEM, rightMargin=MARGEM, topMargin=10 * mm, bottomMargin=10 * mm,
        title=f"DIME Detalhada - {nome} - {periodo}",
    )

    PRETO = colors.black
    AZUL = colors.HexColor('#1155cc')

    # ---- estilos de parágrafo -------------------------------------------------
    st_org = ParagraphStyle('org', fontName='Helvetica', fontSize=10, alignment=TA_CENTER, leading=12)
    st_dime = ParagraphStyle('dime', fontName='Helvetica-Bold', fontSize=13, alignment=TA_CENTER, leading=16)
    st_hdr = ParagraphStyle('hdr', fontName='Helvetica-Bold', fontSize=7.5, alignment=TA_CENTER, leading=9)
    st_hdrL = ParagraphStyle('hdrL', fontName='Helvetica-Bold', fontSize=7.5, alignment=TA_LEFT, leading=9)
    st_cel = ParagraphStyle('cel', fontName='Helvetica', fontSize=7, alignment=TA_LEFT, leading=8.5)
    st_num = ParagraphStyle('num', fontName='Helvetica', fontSize=7, alignment=TA_RIGHT, leading=8.5)
    st_ctr = ParagraphStyle('ctr', fontName='Helvetica', fontSize=7, alignment=TA_CENTER, leading=8.5)
    st_titulo = ParagraphStyle('titulo', fontName='Helvetica-Bold', fontSize=8.5, alignment=TA_LEFT, leading=10)
    st_topo = ParagraphStyle('topo', fontName='Helvetica', fontSize=7.5, alignment=TA_RIGHT,
                              textColor=AZUL, leading=9)
    st_idx = ParagraphStyle('idx', fontName='Helvetica-Bold', fontSize=8.5, alignment=TA_LEFT,
                             textColor=AZUL, leading=13)
    st_anchor = ParagraphStyle('anchor', fontName='Helvetica', fontSize=1, leading=1)

    def P(text, style=st_cel):
        return rl['Paragraph'](str(text), style)

    def ancora(nome_ancora):
        return rl['Paragraph'](f'<a name="{nome_ancora}"/>&nbsp;', st_anchor)

    def topo_link():
        return P('<a href="#topo"><u>Topo</u></a>', st_topo)

    def caixa_titulo(codigo, titulo, ncols, largura_total):
        """Primeira linha de cada quadro: título à esquerda + link Topo à direita,
        dentro da mesma tabela (mesma caixa) do quadro."""
        linha = [P(f"<b>{codigo} - {titulo}</b>", st_titulo)] + [''] * (ncols - 2) + [topo_link()]
        return linha

    def nova_tabela(dados, col_widths, spans=(), header_rows=1, total_rows=()):
        t = rl['Table'](dados, colWidths=col_widths, repeatRows=header_rows + 1)
        estilo = [
            ('GRID', (0, header_rows + 1), (-1, -1), 0.5, PRETO) if header_rows else ('GRID', (0, 0), (-1, -1), 0.5, PRETO),
            ('BOX', (0, 0), (-1, -1), 0.75, PRETO),
            ('LINEBELOW', (0, 0), (-1, 0), 0.75, PRETO),   # sob a barra de título
            ('LINEBELOW', (0, header_rows), (-1, header_rows), 0.75, PRETO),  # sob o(s) cabeçalho(s)
            ('GRID', (0, 1), (-1, header_rows), 0.5, PRETO),               # grade nas linhas de cabeçalho
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 1.6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1.6),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('SPAN', (0, 0), (ncols_menos_um(dados) - 2, 0)),
        ]
        for (c1, r1, c2, r2) in spans:
            estilo.append(('SPAN', (c1, r1), (c2, r2)))
        for r in total_rows:
            estilo.append(('LINEABOVE', (0, r), (-1, r), 0.75, PRETO))
        t.setStyle(rl['TableStyle'](estilo))
        return t

    def ncols_menos_um(dados):
        return len(dados[0])

    story = []

    # ---- cabeçalho --------------------------------------------------------
    story.append(ancora('topo'))
    story.append(rl['Paragraph']("Estado de Santa Catarina", st_org))
    story.append(rl['Paragraph']("Secretaria de Estado da Fazenda", st_org))
    story.append(rl['Paragraph']("Diretoria de Administração Tributária - DIAT", st_org))
    story.append(rl['Paragraph']("DIME", st_dime))
    story.append(rl['Paragraph'](
        "<i>(relatório gerado localmente a partir do arquivo-texto - Anexo I do Manual da DIME - "
        "antes da transmissão à SEF/SC)</i>",
        ParagraphStyle('obs', fontName='Helvetica-Oblique', fontSize=7, alignment=TA_CENTER,
                        textColor=colors.HexColor('#555555'))))
    story.append(rl['Spacer'](1, 6))

    t = rl['Table']([[P('I.E.', st_hdr), P('Contribuinte', st_hdr), P('Período', st_hdr)],
                      [P(ie, st_ctr), P(nome, st_cel), P(periodo, st_ctr)]],
                     colWidths=[LARGURA * 0.20, LARGURA * 0.60, LARGURA * 0.20])
    t.setStyle(rl['TableStyle']([
        ('GRID', (0, 0), (-1, -1), 0.5, PRETO), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(t)
    story.append(rl['Spacer'](1, 8))

    # ---- monta a lista de quadros efetivamente presentes -------------------
    todos_quadros = [('00', 'Informações Iniciais da Declaração', True)]
    todos_quadros.append(('01', 'Valores Fiscais Entradas', bool(decl['quadro01'])))
    todos_quadros.append(('02', 'Valores Fiscais Saídas', bool(decl['quadro02'])))
    for cod in ('03', '04', '05', '09', '10', '11'):
        todos_quadros.append((cod, core.QUADRO_TITULO[cod], bool(decl['quadro' + cod])))
    todos_quadros.append(('12', core.QUADRO_TITULO['12'], bool(decl['quadro12'])))
    todos_quadros.append(('15', core.QUADRO_TITULO['15'], bool(decl['quadro15'])))
    todos_quadros.append(('16', core.QUADRO_TITULO['16'], bool(decl['quadro16'])))
    todos_quadros.append(('46', core.QUADRO_TITULO['46'], bool(decl['quadro46'])))
    todos_quadros.append(('49', core.QUADRO_TITULO['49'], bool(decl['quadro49'])))
    todos_quadros.append(('50', core.QUADRO_TITULO['50'], bool(decl['quadro50'])))
    presentes = [(c, t_) for c, t_, ok in todos_quadros if ok]

    # ---- caixa "QUADROS" (índice com links internos) -----------------------
    linhas_idx = [[P(f'<a href="#q{c}">{c} - {t_}</a>', st_idx)] for c, t_ in presentes]
    tabela_idx = rl['Table'](linhas_idx, colWidths=[LARGURA])
    tabela_idx.setStyle(rl['TableStyle']([
        ('BOX', (0, 0), (-1, -1), 0.75, PRETO),
        ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(rl['Paragraph']("QUADROS", ParagraphStyle('qtit', fontName='Helvetica-Bold', fontSize=10,
                                                             alignment=TA_CENTER, spaceAfter=4)))
    story.append(tabela_idx)
    story.append(rl['Spacer'](1, 10))

    def add_quadro(codigo, tabela_flowable):
        story.append(ancora(f'q{codigo}'))
        story.append(tabela_flowable)
        story.append(rl['Spacer'](1, 8))

    # ---- Quadro 00 ----------------------------------------------------------
    cab = [P('Tipo de declaração', st_hdr), P('Regime de apuração', st_hdr),
           P('Apuração consolidada', st_hdr), P('Apuração centralizada', st_hdr),
           P('Movimento', st_hdr), P('Quantidade de trabalhadores', st_hdr)]
    dados = [P(core.TIPO_DECLARACAO.get(str(q00['tipo_declaracao']), ''), st_ctr),
             P(core.REGIME_APURACAO.get(str(q00['regime_apuracao']), ''), st_ctr),
             P(core.APURACAO_CONSOLIDADA.get(str(q00['apuracao_consolidada']), ''), st_ctr),
             P(core.APURACAO_CENTRALIZADA.get(str(q00['apuracao_centralizada']), ''), st_ctr),
             P(core.MOVIMENTO.get(str(q00['movimento']), ''), st_ctr),
             P(str(q00['qtd_trabalhadores']), st_ctr)]
    n = len(cab)
    linhas = [caixa_titulo('00', 'Informações Iniciais', n, LARGURA), cab, dados]
    larg = LARGURA / n
    add_quadro('00', nova_tabela(linhas, [larg] * n, header_rows=1))

    # ---- Quadro 01 - Entradas (cabeçalho de 2 níveis) -----------------------
    if decl['quadro01']:
        linha1 = [P('CFOP', st_hdr), P('Valor Contábil', st_hdr),
                  P('Operações Com Crédito de Imposto', st_hdr), '',
                  P('Operações Sem Crédito de Imposto', st_hdr), '',
                  P('Débito Substituição Tributária', st_hdr), '',
                  P('Débito Imposto Diferencial de Alíquota', st_hdr)]
        linha2 = ['', '', P('Base de Cálculo', st_hdr), P('Imposto Creditado', st_hdr),
                  P('Isentas/Não Tributadas', st_hdr), P('Outras', st_hdr),
                  P('Base de Cálculo', st_hdr), P('Imposto Retido', st_hdr), '']
        n = 9
        titulo = caixa_titulo('01', 'Valores Fiscais Entradas', n, LARGURA)
        linhas = [titulo, linha1, linha2]
        totais = {c: core.Decimal('0.00') for c in core.CFOP_ENTRADAS_COLS}
        for r in sorted(decl['quadro01'], key=lambda x: x['cfop']):
            linhas.append([P(r['cfop'], st_ctr)] + [P(core.fmt_money(r[c]), st_num) for c in core.CFOP_ENTRADAS_COLS])
            for c in core.CFOP_ENTRADAS_COLS:
                totais[c] += r[c]
        idx_total = len(linhas)
        linhas.append([P('<b>TT</b>', st_ctr)] + [P(f"<b>{core.fmt_money(totais[c])}</b>", st_num) for c in core.CFOP_ENTRADAS_COLS])
        larguras = [LARGURA * w for w in (0.06, 0.135, 0.115, 0.115, 0.125, 0.125, 0.10, 0.10, 0.125)]
        spans = [(2, 1, 3, 1), (4, 1, 5, 1), (6, 1, 7, 1), (0, 1, 0, 2), (1, 1, 1, 2), (8, 1, 8, 2)]
        add_quadro('01', nova_tabela(linhas, larguras, spans=spans, header_rows=2, total_rows=(idx_total,)))

    # ---- Quadro 02 - Saídas (cabeçalho de 2 níveis) --------------------------
    if decl['quadro02']:
        linha1 = [P('CFOP', st_hdr), P('Valor Contábil', st_hdr),
                  P('Operações Com Débito de Imposto', st_hdr), '',
                  P('Operações Sem Débito de Imposto', st_hdr), '',
                  P('Débito Substituição Tributária', st_hdr), '']
        linha2 = ['', '', P('Base de Cálculo', st_hdr), P('Imposto Debitado', st_hdr),
                  P('Isentas/Não Tributadas', st_hdr), P('Outras', st_hdr),
                  P('Base de Cálculo', st_hdr), P('Imposto Retido', st_hdr)]
        n = 8
        titulo = caixa_titulo('02', 'Valores Fiscais Saídas', n, LARGURA)
        linhas = [titulo, linha1, linha2]
        totais = {c: core.Decimal('0.00') for c in core.CFOP_SAIDAS_COLS}
        for r in sorted(decl['quadro02'], key=lambda x: x['cfop']):
            linhas.append([P(r['cfop'], st_ctr)] + [P(core.fmt_money(r[c]), st_num) for c in core.CFOP_SAIDAS_COLS])
            for c in core.CFOP_SAIDAS_COLS:
                totais[c] += r[c]
        idx_total = len(linhas)
        linhas.append([P('<b>TT</b>', st_ctr)] + [P(f"<b>{core.fmt_money(totais[c])}</b>", st_num) for c in core.CFOP_SAIDAS_COLS])
        larguras = [LARGURA * w for w in (0.07, 0.15, 0.13, 0.13, 0.14, 0.14, 0.12, 0.12)]
        spans = [(2, 1, 3, 1), (4, 1, 5, 1), (6, 1, 7, 1), (0, 1, 0, 2), (1, 1, 1, 2)]
        add_quadro('02', nova_tabela(linhas, larguras, spans=spans, header_rows=2, total_rows=(idx_total,)))

    # ---- Quadros-resumo (Item / Descrição / Valor) ---------------------------
    for cod in ('03', '04', '05', '09', '10', '11'):
        itens = decl['quadro' + cod]
        if not itens:
            continue
        n = 3
        titulo = caixa_titulo(cod, core.QUADRO_TITULO[cod], n, LARGURA)
        cab = [P('Item', st_hdr), P('Descrição', st_hdr), P('Valor', st_hdr)]
        linhas = [titulo, cab]
        for item in sorted(itens.keys()):
            codigo = core.display_item_code(cod, item)
            linhas.append([P(codigo, st_ctr), P(core.item_desc(cod, item), st_cel), P(core.fmt_money(itens[item]), st_num)])
        larguras = [LARGURA * 0.08, LARGURA * 0.72, LARGURA * 0.20]
        add_quadro(cod, nova_tabela(linhas, larguras, header_rows=1))

    # ---- Quadro 12 - Pagamentos -----------------------------------------------
    if decl['quadro12']:
        n = 6
        titulo = caixa_titulo('12', core.QUADRO_TITULO['12'], n, LARGURA)
        cab = [P('Origem', st_hdr), P('Código da Receita', st_hdr), P('Classe de Vencimento', st_hdr),
               P('Data de Vencimento', st_hdr), P('Valor', st_hdr), P('Número do Acordo', st_hdr)]
        linhas = [titulo, cab]
        for r in decl['quadro12']:
            linhas.append([P(r['origem'], st_ctr), P(r['codigo_receita'], st_ctr),
                            P(r['classe_vencimento'], st_ctr), P(core.fmt_data(r['data']), st_ctr),
                            P(core.fmt_money(r['valor']), st_num), P(r['numero_acordo'], st_ctr)])
        larguras = [LARGURA * 0.10, LARGURA * 0.18, LARGURA * 0.18, LARGURA * 0.18, LARGURA * 0.18, LARGURA * 0.18]
        add_quadro('12', nova_tabela(linhas, larguras, header_rows=1))

    # ---- Quadro 15 - Fundos ----------------------------------------------------
    if decl['quadro15']:
        n = 8
        titulo = caixa_titulo('15', core.QUADRO_TITULO['15'], n, LARGURA)
        cab = [P('Sequência', st_hdr), P('Código do Benefício TTD', st_hdr), P('Número Concessão TTD', st_hdr),
               P('Subtipo DCIP', st_hdr), P('Valor ICMS Exonerado', st_hdr), P('Valor FUMDES', st_hdr),
               P('Valor FUNDO SOCIAL', st_hdr), P('Valores de Devolução (BC/ICMS/FUMDES/F.SOCIAL)', st_hdr)]
        linhas = [titulo, cab]
        for r in decl['quadro15']:
            devol = (f"{core.fmt_money(r['valor_bc_devolucao'])} / {core.fmt_money(r['valor_icms_exonerado_devolucao'])} / "
                     f"{core.fmt_money(r['valor_fumdes_devolucao'])} / {core.fmt_money(r['valor_fundosocial_devolucao'])}")
            linhas.append([P(r['sequencia'], st_ctr), P(r['codigo_beneficio_ttd'], st_ctr),
                            P(r['numero_concessao'], st_ctr), P(r['subtipo_dcip'], st_ctr),
                            P(core.fmt_money(r['valor_icms_exonerado']), st_num),
                            P(core.fmt_money(r['valor_fumdes']), st_num),
                            P(core.fmt_money(r['valor_fundosocial']), st_num), P(devol, st_num)])
        larguras = [LARGURA * w for w in (0.08, 0.10, 0.14, 0.09, 0.13, 0.11, 0.13, 0.22)]
        add_quadro('15', nova_tabela(linhas, larguras, header_rows=1))

    # ---- Quadro 16 --------------------------------------------------------
    itens = decl['quadro16']
    if itens:
        n = 3
        titulo = caixa_titulo('16', core.QUADRO_TITULO['16'], n, LARGURA)
        cab = [P('Item', st_hdr), P('Descrição', st_hdr), P('Valor', st_hdr)]
        linhas = [titulo, cab]
        for item in sorted(itens.keys()):
            codigo = core.display_item_code('16', item)
            linhas.append([P(codigo, st_ctr), P(core.item_desc('16', item), st_cel), P(core.fmt_money(itens[item]), st_num)])
        larguras = [LARGURA * 0.08, LARGURA * 0.72, LARGURA * 0.20]
        add_quadro('16', nova_tabela(linhas, larguras, header_rows=1))

    # ---- Quadro 46 ----------------------------------------------------------
    if decl['quadro46']:
        n = 4
        titulo = caixa_titulo('46', core.QUADRO_TITULO['46'], n, LARGURA)
        cab = [P('Item', st_hdr), P('Identificação do Regime', st_hdr), P('Valor', st_hdr), P('Origem', st_hdr)]
        linhas = [titulo, cab]
        for r in decl['quadro46']:
            origem_desc = core.ORIGEM_46.get(r['origem'], str(r['origem']))
            linhas.append([P(r['sequencia'], st_ctr), P(f"{r['identificacao']:015d}", st_ctr),
                            P(core.fmt_money(r['valor']), st_num), P(f"{r['origem']} - {origem_desc}", st_cel)])
        larguras = [LARGURA * 0.08, LARGURA * 0.22, LARGURA * 0.18, LARGURA * 0.52]
        add_quadro('46', nova_tabela(linhas, larguras, header_rows=1))

    # ---- Quadro 49 - Entradas por UF ------------------------------------------
    if decl['quadro49']:
        linha1 = [P('UF', st_hdr), P('Valor Contábil', st_hdr), P('Base de Cálculo', st_hdr), P('Outras', st_hdr),
                  P('ICMS Retido por Substituição Tributária', st_hdr), '']
        linha2 = ['', '', '', '', P('Petróleo / Energia Elétrica', st_hdr), P('Outros Produtos', st_hdr)]
        n = 6
        titulo = caixa_titulo('49', core.QUADRO_TITULO['49'], n, LARGURA)
        linhas = [titulo, linha1, linha2]
        mapa = {r['uf']: r for r in decl['quadro49']}
        idx_total = None
        for uf in core.UF_ORDEM:
            if uf in mapa:
                r = mapa[uf]
                if uf == 'TT':
                    idx_total = len(linhas)
                linhas.append([P(uf, st_ctr), P(core.fmt_money(r['valor_contabil']), st_num),
                                P(core.fmt_money(r['base_calculo']), st_num), P(core.fmt_money(r['outras']), st_num),
                                P(core.fmt_money(r['petroleo_energia']), st_num),
                                P(core.fmt_money(r['outros_produtos']), st_num)])
        larguras = [LARGURA * w for w in (0.06, 0.20, 0.20, 0.20, 0.17, 0.17)]
        spans = [(4, 1, 5, 1), (0, 1, 0, 2), (1, 1, 1, 2), (2, 1, 2, 2), (3, 1, 3, 2)]
        add_quadro('49', nova_tabela(linhas, larguras, spans=spans, header_rows=2,
                                      total_rows=(idx_total,) if idx_total else ()))

    # ---- Quadro 50 - Saídas por UF --------------------------------------------
    if decl['quadro50']:
        linha1 = [P('UF', st_hdr), P('Valor Contábil', st_hdr), '', P('Base de Cálculo', st_hdr), '',
                  P('Outras', st_hdr), P('ICMS Retido por Subst. Trib.', st_hdr)]
        linha2 = ['', P('Não Contribuintes', st_hdr), P('Contribuintes', st_hdr),
                  P('Não Contribuintes', st_hdr), P('Contribuintes', st_hdr), '', '']
        n = 7
        titulo = caixa_titulo('50', core.QUADRO_TITULO['50'], n, LARGURA)
        linhas = [titulo, linha1, linha2]
        mapa = {r['uf']: r for r in decl['quadro50']}
        idx_total = None
        for uf in core.UF_ORDEM:
            if uf in mapa:
                r = mapa[uf]
                if uf == 'TT':
                    idx_total = len(linhas)
                linhas.append([P(uf, st_ctr), P(core.fmt_money(r['valor_contabil_nao_contrib']), st_num),
                                P(core.fmt_money(r['valor_contabil_contrib']), st_num),
                                P(core.fmt_money(r['base_calculo_nao_contrib']), st_num),
                                P(core.fmt_money(r['base_calculo_contrib']), st_num),
                                P(core.fmt_money(r['outras']), st_num), P(core.fmt_money(r['icms_st']), st_num)])
        larguras = [LARGURA * w for w in (0.06, 0.16, 0.16, 0.16, 0.16, 0.15, 0.15)]
        spans = [(1, 1, 2, 1), (3, 1, 4, 1), (0, 1, 0, 2), (5, 1, 5, 2), (6, 1, 6, 2)]
        add_quadro('50', nova_tabela(linhas, larguras, spans=spans, header_rows=2,
                                      total_rows=(idx_total,) if idx_total else ()))

    if decl['registros_nao_reconhecidos']:
        story.append(rl['Paragraph'](
            f"Observação: {len(decl['registros_nao_reconhecidos'])} registro(s) de tipo não mapeado por este "
            "conversor foram ignorados.", ParagraphStyle('warn', fontName='Helvetica-Oblique', fontSize=7)))

    doc.build(story)
    log(f"PDF gerado: {caminho_pdf}")
    return caminho_pdf


def converter_arquivo(caminho_txt, pasta_saida, log=print):
    """Processa um arquivo-texto e gera um PDF por declaração encontrada.
    Retorna a lista de caminhos de PDF gerados."""
    contador, declaracoes = core.parse_file(caminho_txt)
    if not declaracoes:
        raise ValueError("Nenhuma declaração (registro tipo 21) foi encontrada no arquivo selecionado.")

    os.makedirs(pasta_saida, exist_ok=True)
    gerados = []
    for decl in declaracoes:
        q00 = decl['quadro00']
        ie_str = f"{q00['inscricao']:09d}"
        periodo_str = str(q00['periodo']).zfill(6)
        nome_arq = f"DIME_DETALHADA_{ie_str}_{periodo_str}.pdf"
        caminho_pdf = os.path.join(pasta_saida, nome_arq)
        log(f"Processando declaração {core.fmt_ie(q00['inscricao'])} - período {core.fmt_periodo(q00['periodo'])} ...")
        gerar_pdf(contador, decl, caminho_pdf, log=log)
        if decl['registros_nao_reconhecidos']:
            log(f"  Aviso: {len(decl['registros_nao_reconhecidos'])} registro(s) de tipo não mapeado foram ignorados.")
        gerados.append(caminho_pdf)
    return gerados


# ---------------------------------------------------------------------------
# Interface gráfica (tkinter)
# ---------------------------------------------------------------------------

def _lancar_gui():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext
    except ImportError as exc:
        raise ImportError(
            "O módulo 'tkinter' não foi encontrado nesta instalação do Python.\n"
            "No Windows, reinstale o Python marcando a opção 'tcl/tk and IDLE' "
            "no instalador oficial (python.org)."
        ) from exc

    class Aplicativo(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("Conversor DIME/GIA-SC → DIME Detalhada (PDF)")
            self.geometry("720x480")
            self.resizable(True, True)

            self.arquivo_entrada = tk.StringVar()
            self.pasta_saida = tk.StringVar()

            pad = {'padx': 10, 'pady': 6}

            frm_topo = tk.Frame(self)
            frm_topo.pack(fill='x', **pad)

            tk.Label(frm_topo, text="Arquivo TXT (DIME/GIA-SC):").grid(row=0, column=0, sticky='w')
            tk.Entry(frm_topo, textvariable=self.arquivo_entrada, width=70).grid(row=1, column=0, sticky='we')
            tk.Button(frm_topo, text="Selecionar...", command=self.selecionar_arquivo).grid(row=1, column=1, padx=6)

            tk.Label(frm_topo, text="Pasta de destino do PDF:").grid(row=2, column=0, sticky='w', pady=(10, 0))
            tk.Entry(frm_topo, textvariable=self.pasta_saida, width=70).grid(row=3, column=0, sticky='we')
            tk.Button(frm_topo, text="Selecionar...", command=self.selecionar_pasta).grid(row=3, column=1, padx=6)

            frm_topo.grid_columnconfigure(0, weight=1)

            self.btn_converter = tk.Button(self, text="Converter para PDF", command=self.converter,
                                            bg="#1c4587", fg="white", height=2)
            self.btn_converter.pack(fill='x', padx=10, pady=10)

            tk.Label(self, text="Log:").pack(anchor='w', padx=10)
            self.log_txt = scrolledtext.ScrolledText(self, height=16)
            self.log_txt.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        def log(self, msg):
            self.log_txt.insert('end', msg + "\n")
            self.log_txt.see('end')
            self.update_idletasks()

        def selecionar_arquivo(self):
            caminho = filedialog.askopenfilename(
                title="Selecione o arquivo-texto DIME/GIA-SC",
                filetypes=[("Arquivo texto", "*.txt"), ("Todos os arquivos", "*.*")])
            if caminho:
                self.arquivo_entrada.set(caminho)
                if not self.pasta_saida.get():
                    self.pasta_saida.set(os.path.dirname(caminho))

        def selecionar_pasta(self):
            pasta = filedialog.askdirectory(title="Selecione a pasta de destino")
            if pasta:
                self.pasta_saida.set(pasta)

        def converter(self):
            entrada = self.arquivo_entrada.get().strip()
            saida = self.pasta_saida.get().strip()
            if not entrada or not os.path.isfile(entrada):
                messagebox.showerror("Erro", "Selecione um arquivo-texto válido.")
                return
            if not saida:
                messagebox.showerror("Erro", "Selecione a pasta de destino.")
                return

            self.btn_converter.config(state='disabled')
            self.log_txt.delete('1.0', 'end')

            def tarefa():
                try:
                    gerados = converter_arquivo(entrada, saida, log=self.log)
                    self.log("")
                    self.log(f"Concluído. {len(gerados)} PDF(s) gerado(s) em: {saida}")
                    messagebox.showinfo("Concluído", f"{len(gerados)} PDF(s) gerado(s) com sucesso em:\n{saida}")
                except Exception:
                    erro = traceback.format_exc()
                    self.log("ERRO:\n" + erro)
                    messagebox.showerror("Erro na conversão", erro.splitlines()[-1])
                finally:
                    self.btn_converter.config(state='normal')

            threading.Thread(target=tarefa, daemon=True).start()

    app = Aplicativo()
    app.mainloop()


def main():
    if len(sys.argv) > 1:
        # Modo linha de comando, sem abrir janela: dime_desktop.py entrada.txt [pasta_saida]
        entrada = sys.argv[1]
        saida = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(entrada)) or '.'
        gerados = converter_arquivo(entrada, saida)
        print(f"{len(gerados)} PDF(s) gerado(s):")
        for g in gerados:
            print(" -", g)
    else:
        _lancar_gui()


if __name__ == '__main__':
    main()
