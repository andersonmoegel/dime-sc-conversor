#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dime_txt_to_detalhada.py
=========================

Converte o arquivo-texto de layout fixo da DIME/GIA-SC (Anexo I da Portaria
SEF/SC - "Manual de Orientação - Preenchimento da DIME") em um relatório no
formato "DIME Detalhada" (o extrato exibido pelo sistema S@T da SEF/SC em
sat.sef.sc.gov.br/tax.net/.../DEC_DIME_Declaracao.aspx).

O layout dos registros (tipos 20, 21, 22, 23, 24, 25, 26, 30, 31, 32, 33,
35, 36, 37, 41, 42, 46, 47, 48, 49, 50, 51, 80-85, 98, 99) foi extraído do
Manual Consolidado da DIME (v31, 22/03/2024), item "5. Layout dos
Registros", e cada campo foi validado byte a byte contra um arquivo real
(GIA_SC ...txt) comparando o resultado com o extrato DIME Detalhada em PDF
correspondente à mesma declaração (WEG TINTAS LTDA, período 05/2026).

USO
---
    python3 dime_txt_to_detalhada.py ARQUIVO.txt [-o PASTA_SAIDA] [--txt]

Para cada declaração de ICMS encontrada no arquivo (delimitada pelos
registros tipo "21" ... "98") é gerado um arquivo HTML
"DIME_DETALHADA_<IE>_<periodo>.html" no formato do relatório oficial.
Com a opção --txt, também é gerado um .txt (relatório em texto simples).

O programa não depende de nenhuma biblioteca externa (apenas Python 3
padrão).
"""

import argparse
import os
import sys
from decimal import Decimal, InvalidOperation
from datetime import datetime

# ---------------------------------------------------------------------------
# 1. LAYOUT DOS REGISTROS (Manual DIME, Anexo I, item 5)
# ---------------------------------------------------------------------------
# Todo registro começa com os campos comuns:
#   posição 001/002 - Tipo de Registro (N)
#   posição 003/004 - Quadro (C ou N, conforme o registro; irrelevante p/ leitura)
# Os campos abaixo de cada tipo são os campos QUE VÊM DEPOIS desses dois.
#
# tipo de campo:
#   'N' inteiro                'C' alfanumérico
#   '$' valor com 2 decimais   'D' data DDMMAAAA
#
# Observação: no Manual, a coluna "Posição" de alguns registros (36, por
# exemplo) contém erros de digitação/OCR; por isso as posições REAIS usadas
# aqui foram recalculadas pela soma cumulativa dos "Tamanhos" (coluna
# confiável) na ordem em que os campos aparecem no Manual, e conferidas
# campo a campo contra um arquivo real.

ITEM_VALOR = [
    ('item', 3, 'C'),      # número do item do quadro (mantém zeros à esquerda)
    ('valor', 17, '$'),
]

LAYOUTS = {
    '20': [                                    # Dados do Contabilista
        ('cpf', 11, 'N'), ('nome', 50, 'C'), ('data_hora', 14, 'N'),
    ],
    '21': [                                    # Quadro 00 - Dados iniciais
        ('inscricao', 9, 'N'), ('nome', 50, 'C'), ('periodo', 6, 'C'),
        ('tipo_declaracao', 1, 'N'), ('regime_apuracao', 1, 'N'), ('porte', 1, 'N'),
        ('apuracao_consolidada', 1, 'N'), ('apuracao_centralizada', 1, 'N'),
        ('transferencia_creditos', 1, 'N'), ('creditos_presumidos', 1, 'N'),
        ('creditos_incentivos', 1, 'N'), ('movimento', 1, 'N'),
        ('substituto_tributario', 1, 'N'), ('escrita_contabil', 1, 'N'),
        ('qtd_trabalhadores', 5, 'N'),
    ],
    '22': [                                    # Quadro 01 - Valores Fiscais Entradas
        ('cfop', 5, 'N'), ('valor_contabil', 17, '$'), ('base_calculo', 17, '$'),
        ('imposto_creditado', 17, '$'), ('isentas', 17, '$'), ('outras', 17, '$'),
        ('bc_imposto_retido', 17, '$'), ('imposto_retido', 17, '$'), ('dif_aliquota', 17, '$'),
    ],
    '23': [                                    # Quadro 02 - Valores Fiscais Saídas
        ('cfop', 5, 'N'), ('valor_contabil', 17, '$'), ('base_calculo', 17, '$'),
        ('imposto_debitado', 17, '$'), ('isentas', 17, '$'), ('outras', 17, '$'),
        ('bc_imposto_retido', 17, '$'), ('imposto_retido', 17, '$'), ('branco', 17, '$'),
    ],
    '24': ITEM_VALOR,                          # Quadro 03 - Resumo dos Valores Fiscais
    '25': ITEM_VALOR,                          # Quadro 04 - Resumo da Apuração dos Débitos
    '26': ITEM_VALOR,                          # Quadro 05 - Resumo da Apuração dos Créditos
    '30': ITEM_VALOR,                          # Quadro 09 - Cálculo do Imposto a Pagar ou Saldo Credor
    '31': ITEM_VALOR,                          # Quadro 10 - Débitos Específicos
    '32': ITEM_VALOR,                          # Quadro 11 - Informações sobre Substituição Tributária
    '33': [                                    # Quadro 12 - Discriminação dos Pagamentos
        ('origem', 1, 'C'), ('codigo_receita', 4, 'N'), ('data', 8, 'D'),
        ('valor', 17, '$'), ('classe_vencimento', 5, 'N'), ('numero_acordo', 15, 'N'),
    ],
    '35': ITEM_VALOR,                          # Quadro 14 - DAICP
    '36': [                                    # Quadro 15 - Valores devidos aos Fundos
        ('sequencia', 3, 'C'), ('codigo_beneficio_ttd', 4, 'C'), ('numero_concessao', 15, 'C'),
        ('subtipo_dcip', 4, 'C'), ('valor_bc', 17, '$'), ('valor_icms_exonerado', 17, '$'),
        ('cod_calc_fumdes', 1, 'C'), ('valor_fumdes', 17, '$'),
        ('cod_calc_fundosocial', 1, 'C'), ('valor_fundosocial', 17, '$'),
        ('valor_bc_devolucao', 17, '$'), ('valor_icms_exonerado_devolucao', 17, '$'),
        ('valor_fumdes_devolucao', 17, '$'), ('valor_fundosocial_devolucao', 17, '$'),
    ],
    '37': ITEM_VALOR,                          # Quadro 16 - Apuração de Valores Devidos/Saldo Credor de Fundos
    '41': ITEM_VALOR,                          # Quadro 41 - Créditos Acumulados
    '42': ITEM_VALOR,                          # Quadro 42 - Débitos por Reserva de Créditos Acumulados
    '46': [                                    # Quadro 46 - Créditos por Regimes e Autorizações Especiais
        ('sequencia', 3, 'C'), ('identificacao', 15, 'N'), ('valor', 17, '$'), ('origem', 2, 'N'),
    ],
    '47': [('codigo_municipio', 5, 'C'), ('valor', 17, '$')],
    '48': [('codigo_municipio', 5, 'C'), ('valor', 17, '$'), ('codigo_tipo_atividade', 3, 'N')],
    '49': [                                    # Quadro 49 - Entradas por Unidade da Federação
        ('uf', 2, 'C'), ('valor_contabil', 17, '$'), ('base_calculo', 17, '$'), ('outras', 17, '$'),
        ('petroleo_energia', 17, '$'), ('outros_produtos', 17, '$'),
    ],
    '50': [                                    # Quadro 50 - Saídas por Unidade da Federação
        ('uf', 2, 'C'), ('valor_contabil_nao_contrib', 17, '$'), ('valor_contabil_contrib', 17, '$'),
        ('base_calculo_nao_contrib', 17, '$'), ('base_calculo_contrib', 17, '$'),
        ('outras', 17, '$'), ('icms_st', 17, '$'),
    ],
    '51': ITEM_VALOR,
    '80': ITEM_VALOR, '81': ITEM_VALOR, '82': ITEM_VALOR,
    '83': ITEM_VALOR, '84': ITEM_VALOR, '85': ITEM_VALOR,
    '98': [('qtd_registros', 5, 'N')],
    '99': [('qtd_registros', 5, 'N'), ('qtd_declaracoes', 5, 'N')],
}

# tipo de registro -> (chave interna, "modo": 'lista' ou 'itens')
LIST_TYPES = {
    '22': 'quadro01', '23': 'quadro02', '33': 'quadro12', '36': 'quadro15',
    '46': 'quadro46', '47': 'quadro47', '48': 'quadro48',
    '49': 'quadro49', '50': 'quadro50',
}
ITEM_TYPES = {
    '24': ('quadro03', '03'), '25': ('quadro04', '04'), '26': ('quadro05', '05'),
    '30': ('quadro09', '09'), '31': ('quadro10', '10'), '32': ('quadro11', '11'),
    '35': ('quadro14', '14'), '37': ('quadro16', '16'), '41': ('quadro41', '41'),
    '42': ('quadro42', '42'), '51': ('quadro51', '51'),
    '80': ('quadro80', '80'), '81': ('quadro81', '81'), '82': ('quadro82', '82'),
    '83': ('quadro83', '83'), '84': ('quadro84', '84'), '85': ('quadro85', '85'),
}

# ---------------------------------------------------------------------------
# 2. TABELAS DE DOMÍNIO / DESCRIÇÕES (Manual DIME)
# ---------------------------------------------------------------------------

UF_NOME = {
    'AC': 'ACRE', 'AL': 'ALAGOAS', 'AP': 'AMAPÁ', 'AM': 'AMAZONAS', 'BA': 'BAHIA',
    'CE': 'CEARÁ', 'DF': 'DISTRITO FEDERAL', 'ES': 'ESPÍRITO SANTO', 'GO': 'GOIÁS',
    'MA': 'MARANHÃO', 'MT': 'MATO GROSSO', 'MS': 'MATO GROSSO DO SUL',
    'MG': 'MINAS GERAIS', 'PA': 'PARÁ', 'PB': 'PARAÍBA', 'PR': 'PARANÁ',
    'PE': 'PERNAMBUCO', 'PI': 'PIAUÍ', 'RN': 'RIO GRANDE DO NORTE',
    'RS': 'RIO GRANDE DO SUL', 'RJ': 'RIO DE JANEIRO', 'RO': 'RONDÔNIA',
    'RR': 'RORAIMA', 'SC': 'SANTA CATARINA', 'SP': 'SÃO PAULO', 'SE': 'SERGIPE',
    'TO': 'TOCANTINS', 'EX': 'EXTERIOR DO PAÍS', 'TT': 'TOTAL (todos os estados)',
}

TIPO_DECLARACAO = {'1': 'NORMAL', '2': 'ENCERRAMENTO DE ATIVIDADES', '4': 'ENQUADRAMENTO NO SIMPLES NACIONAL'}
REGIME_APURACAO = {'2': 'NORMAL', '9': 'PRODUTOR PRIMÁRIO'}
APURACAO_CONSOLIDADA = {
    '1': 'NÃO É APURAÇÃO CONSOLIDADA', '2': 'É ESTABELECIMENTO CONSOLIDADOR',
    '3': 'É ESTABELECIMENTO CONSOLIDADO',
}
APURACAO_CENTRALIZADA = {'1': 'NÃO ESTÁ ENQUADRADA NO SIMPLES OU É ESTABELECIMENTO ÚNICO'}
MOVIMENTO = {'1': 'SEM MOVIMENTO E SEM SALDOS', '2': 'SEM MOVIMENTO E COM SALDOS', '3': 'COM MOVIMENTO'}
SUBST_TRIBUTARIO = {'1': 'SIM', '2': 'NÃO', '3': 'SUBSTITUÍDO SOLIDÁRIO'}
TRANSF_CREDITOS = {
    '1': 'NÃO APUROU OU RESERVOU NEM RECEBEU CRÉDITOS', '2': 'APUROU OU RESERVOU CRÉDITOS',
    '3': 'RECEBEU CRÉDITOS', '4': 'APUROU OU RESERVOU E RECEBEU CRÉDITOS',
    '5': 'APURAÇÃO E RESERVA CRÉDITO SISTEMA COOPERATIVO AGROPECUÁRIO',
}
ESCRITA_CONTABIL = {
    '1': 'SIM, É O ESTABELECIMENTO PRINCIPAL', '2': 'NÃO',
    '3': 'SIM, DADOS INFORMADOS NO ESTABELECIMENTO PRINCIPAL',
}

ORIGEM_PAGAMENTO = {'1': 'Imposto', '2': 'Substituição Tributária', '3': 'Débitos Específicos', '4': 'Fundos'}
ORIGEM_46 = {
    0: 'Origem de créditos por regimes especiais e autorizações especiais (AUC)',
    1: 'Crédito por transferência de créditos', 14: 'Crédito por DCIP',
}

QUADRO_TITULO = {
    '01': 'Valores Fiscais Entradas', '02': 'Valores Fiscais Saídas',
    '03': 'Resumo dos Valores Fiscais', '04': 'Resumo da Apuração dos Débitos',
    '05': 'Resumo da Apuração dos Créditos', '09': 'Cálculo do Imposto a Pagar ou Saldo Credor',
    '10': 'Débitos Específicos (compensáveis ou não após o recolhimento)',
    '11': 'Informações sobre Substituição Tributária',
    '12': 'Discriminação dos Pagamentos do Imposto e dos Débitos Específicos',
    '14': 'Demonstrativo da Apuração do Imposto Devido pela Utilização de Crédito Presumido (DAICP)',
    '15': 'Cálculo dos Valores Devidos aos Fundos e Deduções da Devolução',
    '16': 'Demonstrativo da Apuração de Valores Devidos de Fundos ou Saldo Credor',
    '41': 'Demonstrativo de Créditos Acumulados', '42': 'Débitos por Reserva de Créditos Acumulados',
    '46': 'Créditos por regimes especiais', '49': 'Entradas por Unidade da Federação',
    '50': 'Saídas por Unidade da Federação', '51': 'Exclusões do Valor Adicionado no Mês',
    '80': 'Resumo do Livro Registro de Inventário e Receita Bruta', '81': 'Ativo', '82': 'Passivo',
    '83': 'Demonstração de Resultado', '84': 'Detalhamento das Despesas',
    '85': 'Discriminação das Contribuições ao FIA e FEI',
}

# Descrições oficiais dos itens de cada quadro-resumo (Manual DIME, Anexo I).
QUADRO_ITENS = {
    '03': {
        '010': 'Valor contábil', '020': 'Base de cálculo', '030': 'Imposto creditado',
        '040': 'Operações isentas ou não tributadas', '050': 'Outras operações sem crédito de imposto',
        '053': 'Base de Cálculo Imposto Retido', '054': 'Imposto Retido',
        '057': 'Imposto Diferencial Alíquota',
        '060': 'Valor Contábil', '070': 'Base de Cálculo', '080': 'Imposto debitado',
        '090': 'Operações isentas ou não tributadas', '100': 'Outras operações sem débito de imposto',
        '103': 'Base de Cálculo Imposto Retido', '104': 'Imposto Retido',
    },
    '04': {
        '010': '(+) Débito pelas saídas',
        '020': '(+) Débito por diferencial de alíquota ativo permanente',
        '030': '(+) Débito por diferencial de alíquota de material de uso/consumo',
        '040': '(+) Débito de máquinas/equipamentos importados para ativo permanente',
        '045': '(+) Débito da diferença de alíquota - operação/prestação a consumidor final de outro estado',
        '050': '(+) Estorno de crédito', '060': '(+) Outros estornos de crédito',
        '065': '(+) Estorno de crédito da entrada em decorrência de crédito presumido',
        '070': '(+) Outros débitos', '990': '(=) Subtotal de débitos',
    },
    '05': {
        '010': '(+) Saldo credor do mês anterior', '020': '(+) Crédito pelas entradas',
        '030': '(+) Crédito de ativo permanente',
        '040': '(+) Crédito por diferencial de alíquota material de uso/consumo',
        '045': '(+) Crédito da diferença de alíquota - operação/prestação a consumidor final de outro estado',
        '050': '(+) Crédito de ICMS retido por substituição tributária',
        '060': 'Estorno de débito por transferência de créditos acumulados',
        '070': 'Outros estornos de débitos', '080': 'Total de créditos presumidos',
        '090': 'Total de créditos por incentivos fiscais',
        '100': 'Crédito relativo a operações de importação',
        '110': 'Crédito relativo à aquisição de atacadistas de outras UF',
        '120': 'Créditos por responsabilidade tributária',
        '130': 'Outros créditos de pagamentos devidos por ocasião do fato gerador',
        '140': 'Total de créditos por regime especial', '150': 'Restituição de ICMS',
        '160': 'Outros créditos', '990': '(=) Subtotal de créditos',
    },
    '09': {
        '010': '(+) Subtotal de débitos',
        '011': '(+) Complemento de débito por mudança de regime de apuração',
        '020': '(+) Saldos devedores recebidos de estabelecimentos consolidados',
        '030': '(+) Débito por reserva de crédito acumulado',
        '036': '(+) Segregação do crédito presumido utilizado em substituição aos créditos pelas entradas',
        '037': '(+) Segregação do crédito decorrente do pagamento antecipado do ICMS na saída subsequente à importação',
        '038': '(+) Segregação de outros créditos permitidos para compensar com débito por crédito presumido',
        '040': '(=) Total de débitos', '050': '(+) Subtotal de créditos',
        '060': '(+) Saldos credores recebidos de estabelecimentos consolidados',
        '070': '(+) Créditos recebidos por transferência de outros contribuintes',
        '075': '(+) Créditos DCIP',
        '076': '(+) Segregação dos débitos das saídas com crédito presumido em substituição aos créditos pelas entradas',
        '080': '(=) Total de créditos', '090': '(+) Imposto do 1º decêndio',
        '100': '(+) Imposto do 2º decêndio', '105': '(+) Antecipações combustíveis líquidos e gasosos',
        '110': '(=) Total de ajustes da apuração decendial e antecipações',
        '120': '(=) Saldo devedor (Total Débitos - Total Créditos - Total Ajustes Ap Decendial)',
        '130': '(-) Saldo devedor transferido ao estabelecimento consolidador',
        '999': '(=) Imposto a recolher',
        '140': '(=) Saldo credor', '150': '(-) Saldo credor transferido ao estabelecimento consolidador',
        '998': '(=) Saldo Credor para o Mês Seguinte',
        '160': 'Saldo credor transferível relativo à exportação',
        '170': 'Saldo credor transferível relativo a saídas isentas',
        '180': 'Saldo credor transferível relativo a saídas diferidas',
        '190': 'Saldo credor relativo a outros créditos',
    },
    '10': {
        '010': '(+) Débito relativo a operações de importação',
        '020': '(+) Débito relativo à entrada de mercadorias de outras Unidades da Federação',
        '030': '(+) Débito por responsabilidade tributária',
        '040': '(+) Débito por ocasião do fato gerador relativo à saída da mercadoria ou prestação de serviço',
        '050': '(+) Outros débitos eventuais', '990': '(=) Total de débitos',
    },
    '11': {
        '010': 'Valor dos produtos', '020': 'Valor do IPI', '030': 'Despesas acessórias',
        '040': 'Base de cálculo do ICMS próprio', '050': 'ICMS próprio',
        '060': 'Base cálculo ICMS substituição tributária',
        '065': 'Imposto Retido apurado por mercadoria e recolhido por operação',
        '070': '(+) Total do Imposto Retido',
        '073': '(+) Imposto ST devido na entrada com regime especial para apuração mensal',
        '074': '(+) Outros Débitos', '075': '(+) Saldos devedores recebidos de estab. consolidados',
        '080': 'Total de débitos', '090': '(+) Saldo credor do período anterior sobre a ST',
        '100': '(+) Devolução de mercadorias e desfazimento de venda',
        '105': '(+) Créditos declarados no DCIP', '110': '(+) Ressarcimento de ICMS ST',
        '115': '(+) Ressarcimento do ICMS ST acobertado por NF-e',
        '116': '(+) Devolução de mercadorias e desfazimento de venda', '120': '(+) Outros créditos',
        '125': '(+) Saldos credores recebidos de estab. consolidados', '130': '(=) Total de créditos',
        '140': 'Imposto do primeiro decêndio', '150': 'Imposto do segundo decêndio',
        '155': '(+) Antecipações combustíveis líquidos e gasosos',
        '160': '(=) Total de ajustes das antecipações combustíveis', '170': '(=) Saldo devedor',
        '180': '(-) Saldo devedor transferido ao estabelecimento consolidador',
        '999': '(=) Imposto a recolher sobre a substituição tributária', '190': '(=) Saldo credor',
        '200': '(-) Saldo credor transferido ao estabelecimento consolidador',
        '998': '(=) Saldo Credor para o mês seguinte sobre a ST',
        '899': '(=) Débitos Especiais de Substituição Tributária',
    },
    '16': {
        '010': '(+) Soma valores devidos aos FUNDES', '020': '(-) Saldo Credor Mês Anterior FUNDES',
        '030': '(+) Soma Valores FUNDES Devolução',
        '098': '(=) Saldo Credor para o Mês Seguinte FUMDES',
        '099': '(=) FUNDES a Recolher',
        '110': '(+) Soma valores devidos aos FUNDOSOCIAL',
        '120': '(-) Saldo Credor Mês Anterior FUNDOSOCIAL',
        '130': '(+) Soma Valores FUNDOSOCIAL Devolução',
        '198': '(=) Saldo Credor para o Mês Seguinte FUNDOSOCIAL',
        '199': '(=) FUNDOSOCIAL a Recolher',
    },
}

CFOP_ENTRADAS_COLS = ['valor_contabil', 'base_calculo', 'imposto_creditado', 'isentas', 'outras',
                       'bc_imposto_retido', 'imposto_retido', 'dif_aliquota']
CFOP_SAIDAS_COLS = ['valor_contabil', 'base_calculo', 'imposto_debitado', 'isentas', 'outras',
                     'bc_imposto_retido', 'imposto_retido']

UF_ORDEM = ['AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG', 'PA',
            'PB', 'PR', 'PE', 'PI', 'RN', 'RS', 'RJ', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO', 'EX', 'TT']

# ---------------------------------------------------------------------------
# 3. LEITURA / CONVERSÃO DE CAMPOS
# ---------------------------------------------------------------------------


def _to_decimal(raw):
    raw = raw.strip()
    if raw == '':
        return Decimal('0.00')
    try:
        return (Decimal(raw) / Decimal(100)).quantize(Decimal('0.01'))
    except InvalidOperation:
        return Decimal('0.00')


def _to_int(raw):
    raw = raw.strip()
    if raw == '':
        return 0
    try:
        return int(raw)
    except ValueError:
        return 0


def _to_date(raw):
    raw = raw.strip()
    if len(raw) != 8 or raw == '0' * 8:
        return None
    try:
        return datetime.strptime(raw, '%d%m%Y').date()
    except ValueError:
        return None


def convert(raw, ftype):
    if ftype == 'N':
        return _to_int(raw)
    if ftype == '$':
        return _to_decimal(raw)
    if ftype == 'D':
        return _to_date(raw)
    return raw.strip()  # 'C'


def parse_record(line):
    """Decodifica uma linha (registro) segundo o LAYOUTS. Retorna dict."""
    tipo = line[0:2]
    quadro_raw = line[2:4]
    rec = {'_tipo': tipo, '_quadro_raw': quadro_raw}
    layout = LAYOUTS.get(tipo)
    if layout is None:
        rec['_raw'] = line
        return rec
    pos = 4
    for name, length, ftype in layout:
        raw = line[pos:pos + length]
        pos += length
        rec[name] = convert(raw, ftype)
    return rec


# ---------------------------------------------------------------------------
# 4. MONTAGEM DAS DECLARAÇÕES
# ---------------------------------------------------------------------------

def novo_bloco_declaracao():
    bloco = {
        'quadro00': None,
        'quadro01': [], 'quadro02': [],
        'quadro03': {}, 'quadro04': {}, 'quadro05': {}, 'quadro09': {}, 'quadro10': {}, 'quadro11': {},
        'quadro12': [], 'quadro14': {}, 'quadro15': [], 'quadro16': {}, 'quadro41': {}, 'quadro42': {},
        'quadro46': [], 'quadro47': [], 'quadro48': [], 'quadro49': [], 'quadro50': [], 'quadro51': {},
        'quadro80': {}, 'quadro81': {}, 'quadro82': {}, 'quadro83': {}, 'quadro84': {}, 'quadro85': {},
        'registros_nao_reconhecidos': [],
    }
    return bloco


def parse_file(path):
    """Lê o arquivo texto e devolve (contador, [declaracoes])."""
    raw_bytes = open(path, 'rb').read()
    try:
        text = raw_bytes.decode('utf-8')
    except UnicodeDecodeError:
        text = raw_bytes.decode('latin-1')

    linhas = [l for l in text.splitlines() if l.strip() != '']

    contador = None
    declaracoes = []
    atual = None

    for linha in linhas:
        if len(linha) < 4:
            continue
        rec = parse_record(linha)
        tipo = rec['_tipo']

        if tipo == '20':
            contador = rec
            continue
        if tipo == '21':
            atual = novo_bloco_declaracao()
            atual['quadro00'] = rec
            continue
        if tipo == '98':
            if atual is not None:
                atual['_qtd_registros_declaradas'] = rec.get('qtd_registros')
                declaracoes.append(atual)
                atual = None
            continue
        if tipo == '99':
            continue  # fechamento do arquivo - não pertence a nenhuma declaração

        if atual is None:
            continue  # registro fora do escopo de uma declaração (ignorado)

        if tipo in LIST_TYPES:
            atual[LIST_TYPES[tipo]].append(rec)
        elif tipo in ITEM_TYPES:
            chave, _quadro = ITEM_TYPES[tipo]
            atual[chave][rec['item']] = rec['valor']
        else:
            atual['registros_nao_reconhecidos'].append(linha)

    if atual is not None:
        # arquivo sem registro 98 de fechamento
        declaracoes.append(atual)

    return contador, declaracoes


# ---------------------------------------------------------------------------
# 5. FORMATAÇÃO
# ---------------------------------------------------------------------------

def fmt_money(v):
    if v is None:
        v = Decimal('0.00')
    s = f"{v:,.2f}"
    return s.replace(',', '§').replace('.', ',').replace('§', '.')


def fmt_ie(n):
    s = f"{n:09d}"
    return f"{s[0:2]}.{s[2:5]}.{s[5:8]}-{s[8]}"


def fmt_periodo(raw):
    raw = str(raw).zfill(6)
    return f"{raw[0:2]}/{raw[2:6]}"


def fmt_data(d):
    if d is None:
        return ''
    return d.strftime('%d/%m/%Y')


def display_item_code(quadro, item):
    """Reproduz o código exibido no relatório oficial (ex.: quadro 09, item
    999 => '9999'; quadro 10, item 010 => '10010')."""
    return f"{int(quadro)}{item}"


def item_desc(quadro, item):
    return QUADRO_ITENS.get(quadro, {}).get(item, f"Item {item}")


# ---------------------------------------------------------------------------
# 6. GERAÇÃO DO RELATÓRIO HTML (formato "DIME Detalhada")
# ---------------------------------------------------------------------------

CSS = """
body { font-family: Arial, Helvetica, sans-serif; font-size: 12px; color:#222; margin: 24px; }
h1 { font-size: 16px; text-align:center; margin: 2px 0; }
h2 { font-size: 13px; text-align:center; margin: 2px 0; }
.cabecalho { text-align:center; margin-bottom: 18px; }
table { border-collapse: collapse; width: 100%; margin: 6px 0 22px 0; }
th, td { border: 1px solid #999; padding: 3px 6px; font-size: 11px; }
th { background:#e8e8e8; text-align:center; }
td.num { text-align:right; white-space:nowrap; }
td.center { text-align:center; }
.quadro-titulo { background:#cfe2f3; font-weight:bold; padding:5px 8px; margin-top:18px;
                  border:1px solid #999; font-size:12.5px; }
.topo { float:right; font-size:10px; }
.total-row td { font-weight:bold; background:#f2f2f2; }
.vazio { color:#777; font-style: italic; padding: 4px 2px; }
"""


def _linha_vazia(colspan):
    return f'<tr><td class="vazio" colspan="{colspan}">(sem lançamentos)</td></tr>'


def gerar_html(contador, decl):
    q00 = decl['quadro00']
    ie = fmt_ie(q00['inscricao'])
    nome = q00['nome']
    periodo = fmt_periodo(q00['periodo'])

    partes = []
    partes.append(f"<!DOCTYPE html><html lang='pt-BR'><head><meta charset='utf-8'>")
    partes.append(f"<title>DIME Detalhada - {nome} - {periodo}</title><style>{CSS}</style></head><body>")
    partes.append("<div class='cabecalho'>")
    partes.append("<h1>Estado de Santa Catarina — Secretaria de Estado da Fazenda</h1>")
    partes.append("<h2>DIME — Declaração de Informações do ICMS e Movimento Econômico (gerado a partir do arquivo-texto)</h2>")
    partes.append("</div>")

    partes.append("<table><tr><th>I.E.</th><th>Contribuinte</th><th>Período</th></tr>")
    partes.append(f"<tr><td class='center'>{ie}</td><td>{nome}</td><td class='center'>{periodo}</td></tr></table>")

    # Quadro 00
    partes.append("<div class='quadro-titulo'>00 - Informações Iniciais da Declaração</div>")
    partes.append("<table><tr><th>Tipo de declaração</th><th>Regime de apuração</th>"
                   "<th>Apuração consolidada</th><th>Apuração centralizada</th>"
                   "<th>Transferência de créditos</th><th>Movimento</th>"
                   "<th>Substituto tributário</th><th>Escrita contábil</th>"
                   "<th>Qtde. de trabalhadores</th></tr>")
    partes.append("<tr>"
                   f"<td class='center'>{TIPO_DECLARACAO.get(str(q00['tipo_declaracao']), q00['tipo_declaracao'])}</td>"
                   f"<td class='center'>{REGIME_APURACAO.get(str(q00['regime_apuracao']), q00['regime_apuracao'])}</td>"
                   f"<td class='center'>{APURACAO_CONSOLIDADA.get(str(q00['apuracao_consolidada']), q00['apuracao_consolidada'])}</td>"
                   f"<td class='center'>{APURACAO_CENTRALIZADA.get(str(q00['apuracao_centralizada']), q00['apuracao_centralizada'])}</td>"
                   f"<td class='center'>{TRANSF_CREDITOS.get(str(q00['transferencia_creditos']), q00['transferencia_creditos'])}</td>"
                   f"<td class='center'>{MOVIMENTO.get(str(q00['movimento']), q00['movimento'])}</td>"
                   f"<td class='center'>{SUBST_TRIBUTARIO.get(str(q00['substituto_tributario']), q00['substituto_tributario'])}</td>"
                   f"<td class='center'>{ESCRITA_CONTABIL.get(str(q00['escrita_contabil']), q00['escrita_contabil'])}</td>"
                   f"<td class='center'>{q00['qtd_trabalhadores']}</td>"
                   "</tr></table>")

    # Quadro 01 - Entradas
    partes.append("<div class='quadro-titulo'>01 - Valores Fiscais Entradas</div>")
    partes.append("<table><tr><th>CFOP</th><th>Valor Contábil</th><th>Base de Cálculo (c/créd.)</th>"
                   "<th>Imposto Creditado</th><th>Isentas/Não Tributadas</th><th>Outras</th>"
                   "<th>Base Cálc. Subst. Trib.</th><th>Imposto Retido</th><th>Diferencial Alíquota</th></tr>")
    if decl['quadro01']:
        totais = {c: Decimal('0.00') for c in CFOP_ENTRADAS_COLS}
        for r in sorted(decl['quadro01'], key=lambda x: x['cfop']):
            partes.append("<tr><td class='center'>" + str(r['cfop']) + "</td>" +
                           "".join(f"<td class='num'>{fmt_money(r[c])}</td>" for c in CFOP_ENTRADAS_COLS) +
                           "</tr>")
            for c in CFOP_ENTRADAS_COLS:
                totais[c] += r[c]
        partes.append("<tr class='total-row'><td class='center'>TOTAL</td>" +
                       "".join(f"<td class='num'>{fmt_money(totais[c])}</td>" for c in CFOP_ENTRADAS_COLS) + "</tr>")
    else:
        partes.append(_linha_vazia(9))
    partes.append("</table>")

    # Quadro 02 - Saídas
    partes.append("<div class='quadro-titulo'>02 - Valores Fiscais Saídas</div>")
    partes.append("<table><tr><th>CFOP</th><th>Valor Contábil</th><th>Base de Cálculo (c/déb.)</th>"
                   "<th>Imposto Debitado</th><th>Isentas/Não Tributadas</th><th>Outras</th>"
                   "<th>Base Cálc. Subst. Trib.</th><th>Imposto Retido</th></tr>")
    if decl['quadro02']:
        totais = {c: Decimal('0.00') for c in CFOP_SAIDAS_COLS}
        for r in sorted(decl['quadro02'], key=lambda x: x['cfop']):
            partes.append("<tr><td class='center'>" + str(r['cfop']) + "</td>" +
                           "".join(f"<td class='num'>{fmt_money(r[c])}</td>" for c in CFOP_SAIDAS_COLS) +
                           "</tr>")
            for c in CFOP_SAIDAS_COLS:
                totais[c] += r[c]
        partes.append("<tr class='total-row'><td class='center'>TOTAL</td>" +
                       "".join(f"<td class='num'>{fmt_money(totais[c])}</td>" for c in CFOP_SAIDAS_COLS) + "</tr>")
    else:
        partes.append(_linha_vazia(8))
    partes.append("</table>")

    # Quadros-resumo genéricos (item/valor)
    for quadro in ['03', '04', '05', '09', '10', '11']:
        chave = 'quadro' + quadro
        itens = decl[chave]
        partes.append(f"<div class='quadro-titulo'>{quadro} - {QUADRO_TITULO[quadro]}</div>")
        partes.append("<table><tr><th style='width:70px'>Item</th><th>Descrição</th><th style='width:160px'>Valor</th></tr>")
        if itens:
            for item in sorted(itens.keys()):
                cod = display_item_code(quadro, item)
                partes.append(f"<tr><td class='center'>{cod}</td><td>{item_desc(quadro, item)}</td>"
                               f"<td class='num'>{fmt_money(itens[item])}</td></tr>")
        else:
            partes.append(_linha_vazia(3))
        partes.append("</table>")

    # Quadro 12 - Pagamentos
    partes.append("<div class='quadro-titulo'>12 - Discriminação dos Pagamentos do Imposto e dos Débitos Específicos</div>")
    partes.append("<table><tr><th>Origem</th><th>Código da Receita</th><th>Classe de Vencimento</th>"
                   "<th>Data de Vencimento</th><th>Valor</th><th>Número do Acordo</th></tr>")
    if decl['quadro12']:
        for r in decl['quadro12']:
            partes.append("<tr>"
                           f"<td class='center'>{r['origem']} - {ORIGEM_PAGAMENTO.get(str(r['origem']), '')}</td>"
                           f"<td class='center'>{r['codigo_receita']}</td>"
                           f"<td class='center'>{r['classe_vencimento']}</td>"
                           f"<td class='center'>{fmt_data(r['data'])}</td>"
                           f"<td class='num'>{fmt_money(r['valor'])}</td>"
                           f"<td class='center'>{r['numero_acordo']}</td>"
                           "</tr>")
    else:
        partes.append(_linha_vazia(6))
    partes.append("</table>")

    # Quadro 15 - Fundos
    partes.append("<div class='quadro-titulo'>15 - Cálculo dos Valores Devidos aos Fundos e Deduções da Devolução</div>")
    partes.append("<table><tr><th>Seq.</th><th>Cód. TTD</th><th>Nº Concessão</th><th>Subtipo DCIP</th>"
                   "<th>BC ICMS Concessão</th><th>ICMS Exonerado</th><th>FUMDES</th><th>FUNDO SOCIAL</th>"
                   "<th>BC Devolução</th><th>ICMS Exon. Devol.</th><th>FUMDES Devol.</th><th>FUNDO SOCIAL Devol.</th></tr>")
    if decl['quadro15']:
        for r in decl['quadro15']:
            partes.append("<tr>"
                           f"<td class='center'>{r['sequencia']}</td>"
                           f"<td class='center'>{r['codigo_beneficio_ttd']}</td>"
                           f"<td class='center'>{r['numero_concessao']}</td>"
                           f"<td class='center'>{r['subtipo_dcip']}</td>"
                           f"<td class='num'>{fmt_money(r['valor_bc'])}</td>"
                           f"<td class='num'>{fmt_money(r['valor_icms_exonerado'])}</td>"
                           f"<td class='num'>{fmt_money(r['valor_fumdes'])}</td>"
                           f"<td class='num'>{fmt_money(r['valor_fundosocial'])}</td>"
                           f"<td class='num'>{fmt_money(r['valor_bc_devolucao'])}</td>"
                           f"<td class='num'>{fmt_money(r['valor_icms_exonerado_devolucao'])}</td>"
                           f"<td class='num'>{fmt_money(r['valor_fumdes_devolucao'])}</td>"
                           f"<td class='num'>{fmt_money(r['valor_fundosocial_devolucao'])}</td>"
                           "</tr>")
    else:
        partes.append(_linha_vazia(12))
    partes.append("</table>")

    # Quadro 16
    chave = 'quadro16'
    itens = decl[chave]
    partes.append("<div class='quadro-titulo'>16 - Demonstrativo da Apuração de Valores Devidos de Fundos ou Saldo Credor</div>")
    partes.append("<table><tr><th style='width:70px'>Item</th><th>Descrição</th><th style='width:160px'>Valor</th></tr>")
    if itens:
        for item in sorted(itens.keys()):
            cod = display_item_code('16', item)
            partes.append(f"<tr><td class='center'>{cod}</td><td>{item_desc('16', item)}</td>"
                           f"<td class='num'>{fmt_money(itens[item])}</td></tr>")
    else:
        partes.append(_linha_vazia(3))
    partes.append("</table>")

    # Quadro 46
    partes.append("<div class='quadro-titulo'>46 - Créditos por regimes especiais</div>")
    partes.append("<table><tr><th>Item</th><th>Identificação do Regime</th><th>Valor</th><th>Origem</th></tr>")
    if decl['quadro46']:
        for r in decl['quadro46']:
            origem_desc = ORIGEM_46.get(r['origem'], str(r['origem']))
            partes.append("<tr>"
                           f"<td class='center'>{r['sequencia']}</td>"
                           f"<td class='center'>{r['identificacao']:015d}</td>"
                           f"<td class='num'>{fmt_money(r['valor'])}</td>"
                           f"<td class='center'>{r['origem']} - {origem_desc}</td>"
                           "</tr>")
    else:
        partes.append(_linha_vazia(4))
    partes.append("</table>")

    # Quadro 49 - Entradas por UF
    partes.append("<div class='quadro-titulo'>49 - Entradas por Unidade da Federação</div>")
    partes.append("<table><tr><th>UF</th><th>Valor Contábil</th><th>Base de Cálculo</th><th>Outras</th>"
                   "<th>ST - Petróleo/Energia</th><th>ST - Outros Produtos</th></tr>")
    if decl['quadro49']:
        mapa = {r['uf']: r for r in decl['quadro49']}
        for uf in UF_ORDEM:
            if uf in mapa:
                r = mapa[uf]
                cls = 'total-row' if uf == 'TT' else ''
                partes.append(f"<tr class='{cls}'><td class='center'>{uf}</td>"
                               f"<td class='num'>{fmt_money(r['valor_contabil'])}</td>"
                               f"<td class='num'>{fmt_money(r['base_calculo'])}</td>"
                               f"<td class='num'>{fmt_money(r['outras'])}</td>"
                               f"<td class='num'>{fmt_money(r['petroleo_energia'])}</td>"
                               f"<td class='num'>{fmt_money(r['outros_produtos'])}</td></tr>")
    else:
        partes.append(_linha_vazia(6))
    partes.append("</table>")

    # Quadro 50 - Saídas por UF
    partes.append("<div class='quadro-titulo'>50 - Saídas por Unidade da Federação</div>")
    partes.append("<table><tr><th>UF</th><th>Valor Contábil Não Contrib.</th><th>Valor Contábil Contrib.</th>"
                   "<th>Base Cálculo Não Contrib.</th><th>Base Cálculo Contrib.</th><th>Outras</th>"
                   "<th>ICMS Retido por Subst. Trib.</th></tr>")
    if decl['quadro50']:
        mapa = {r['uf']: r for r in decl['quadro50']}
        for uf in UF_ORDEM:
            if uf in mapa:
                r = mapa[uf]
                cls = 'total-row' if uf == 'TT' else ''
                partes.append(f"<tr class='{cls}'><td class='center'>{uf}</td>"
                               f"<td class='num'>{fmt_money(r['valor_contabil_nao_contrib'])}</td>"
                               f"<td class='num'>{fmt_money(r['valor_contabil_contrib'])}</td>"
                               f"<td class='num'>{fmt_money(r['base_calculo_nao_contrib'])}</td>"
                               f"<td class='num'>{fmt_money(r['base_calculo_contrib'])}</td>"
                               f"<td class='num'>{fmt_money(r['outras'])}</td>"
                               f"<td class='num'>{fmt_money(r['icms_st'])}</td></tr>")
    else:
        partes.append(_linha_vazia(7))
    partes.append("</table>")

    if decl['registros_nao_reconhecidos']:
        partes.append("<div class='quadro-titulo'>Registros não mapeados neste conversor</div>")
        partes.append("<table><tr><th>Linha</th></tr>")
        for l in decl['registros_nao_reconhecidos']:
            partes.append(f"<tr><td>{l}</td></tr>")
        partes.append("</table>")

    partes.append("</body></html>")
    return "\n".join(partes)


# ---------------------------------------------------------------------------
# 7. GERAÇÃO DO RELATÓRIO EM TEXTO SIMPLES
# ---------------------------------------------------------------------------

def gerar_texto(contador, decl):
    q00 = decl['quadro00']
    linhas = []
    linhas.append("=" * 100)
    linhas.append("DIME DETALHADA (gerado a partir do arquivo-texto)".center(100))
    linhas.append("=" * 100)
    linhas.append(f"I.E.: {fmt_ie(q00['inscricao'])}    Contribuinte: {q00['nome']}    "
                   f"Período: {fmt_periodo(q00['periodo'])}")
    linhas.append("")

    linhas.append("01 - VALORES FISCAIS ENTRADAS")
    linhas.append("-" * 100)
    for r in sorted(decl['quadro01'], key=lambda x: x['cfop']):
        linhas.append(f"CFOP {r['cfop']:<6} Vlr.Cont.:{fmt_money(r['valor_contabil']):>16} "
                       f"BC:{fmt_money(r['base_calculo']):>16} Imp.Cred.:{fmt_money(r['imposto_creditado']):>16} "
                       f"Isentas:{fmt_money(r['isentas']):>14} Outras:{fmt_money(r['outras']):>14}")
    linhas.append("")

    linhas.append("02 - VALORES FISCAIS SAÍDAS")
    linhas.append("-" * 100)
    for r in sorted(decl['quadro02'], key=lambda x: x['cfop']):
        linhas.append(f"CFOP {r['cfop']:<6} Vlr.Cont.:{fmt_money(r['valor_contabil']):>16} "
                       f"BC:{fmt_money(r['base_calculo']):>16} Imp.Deb.:{fmt_money(r['imposto_debitado']):>16} "
                       f"Isentas:{fmt_money(r['isentas']):>14} Outras:{fmt_money(r['outras']):>14}")
    linhas.append("")

    for quadro in ['03', '04', '05', '09', '10', '11', '16']:
        chave = 'quadro' + quadro
        itens = decl[chave]
        if not itens:
            continue
        linhas.append(f"{quadro} - {QUADRO_TITULO[quadro]}")
        linhas.append("-" * 100)
        for item in sorted(itens.keys()):
            cod = display_item_code(quadro, item)
            linhas.append(f"{cod:<8} {item_desc(quadro, item):<80} {fmt_money(itens[item]):>16}")
        linhas.append("")

    if decl['quadro12']:
        linhas.append("12 - DISCRIMINAÇÃO DOS PAGAMENTOS")
        linhas.append("-" * 100)
        for r in decl['quadro12']:
            linhas.append(f"Origem {r['origem']} ({ORIGEM_PAGAMENTO.get(str(r['origem']), '')})  "
                           f"Receita {r['codigo_receita']}  Classe {r['classe_vencimento']}  "
                           f"Venc. {fmt_data(r['data'])}  Valor {fmt_money(r['valor']):>16}")
        linhas.append("")

    if decl['quadro46']:
        linhas.append("46 - CRÉDITOS POR REGIMES ESPECIAIS")
        linhas.append("-" * 100)
        for r in decl['quadro46']:
            linhas.append(f"Seq {r['sequencia']}  Identificação {r['identificacao']:015d}  "
                           f"Valor {fmt_money(r['valor']):>16}  Origem {r['origem']}")
        linhas.append("")

    for quadro, chave in (('49', 'quadro49'), ('50', 'quadro50')):
        if decl[chave]:
            linhas.append(f"{quadro} - {QUADRO_TITULO[quadro]}")
            linhas.append("-" * 100)
            mapa = {r['uf']: r for r in decl[chave]}
            for uf in UF_ORDEM:
                if uf in mapa:
                    r = mapa[uf]
                    vals = "  ".join(fmt_money(v) for k, v in r.items() if k != 'uf' and k != '_tipo' and k != '_quadro_raw')
                    linhas.append(f"{uf:<4} {vals}")
            linhas.append("")

    return "\n".join(linhas)


# ---------------------------------------------------------------------------
# 8. PROGRAMA PRINCIPAL
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Converte arquivo-texto DIME/GIA-SC em relatório DIME Detalhada.")
    ap.add_argument('arquivo', help='Arquivo texto de entrada (layout fixo DIME/GIA-SC)')
    ap.add_argument('-o', '--saida', default='.', help='Pasta de saída (padrão: pasta atual)')
    ap.add_argument('--txt', action='store_true', help='Também gera versão em texto simples (.txt)')
    args = ap.parse_args()

    os.makedirs(args.saida, exist_ok=True)
    contador, declaracoes = parse_file(args.arquivo)

    if not declaracoes:
        print("Nenhuma declaração (registro tipo 21) encontrada no arquivo.", file=sys.stderr)
        sys.exit(1)

    gerados = []
    for decl in declaracoes:
        q00 = decl['quadro00']
        ie_str = f"{q00['inscricao']:09d}"
        periodo_str = str(q00['periodo']).zfill(6)
        base_nome = f"DIME_DETALHADA_{ie_str}_{periodo_str}"

        html = gerar_html(contador, decl)
        caminho_html = os.path.join(args.saida, base_nome + ".html")
        with open(caminho_html, 'w', encoding='utf-8') as f:
            f.write(html)
        gerados.append(caminho_html)

        if args.txt:
            texto = gerar_texto(contador, decl)
            caminho_txt = os.path.join(args.saida, base_nome + ".txt")
            with open(caminho_txt, 'w', encoding='utf-8') as f:
                f.write(texto)
            gerados.append(caminho_txt)

    print(f"{len(declaracoes)} declaração(ões) processada(s). Arquivos gerados:")
    for g in gerados:
        print(" -", g)


if __name__ == '__main__':
    main()
