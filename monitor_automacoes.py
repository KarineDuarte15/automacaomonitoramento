#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitoramento OptimusBI - Completo:
- Verifica pastas e logs a cada 60s
- A cada 30 minutos: gera o .xlsx em C:\\Users\\agenor.sales\\Desktop\\dashboard acompanhamento
  e faz o upload automático em https://acompanhameto-fluxo-diario-de-taref.vercel.app/ na aba Monitoramento
- A cada 6 horas (360 min): envia resumo no WhatsApp (1 mensagem por contato, fechando e reiniciando o navegador)
- Às 18h: guarda também o relatório diário final na pasta G:\\...
"""
import argparse
import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from datetime import datetime, date, time as dtime, timedelta
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import quote

# ==============================================================================
# 1. CAMINHOS BASE, ARQUIVOS DE ESTADO E LOGGING
# ==============================================================================
AQUI = Path(__file__).resolve().parent
CFG = AQUI / 'config_monitores.json'
HIST = AQUI / 'historico.json'
ESTADO = AQUI / 'estado.json'
SNAPS = AQUI / 'snapshots_hora.json'

MES = [
    'JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
    'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'
]

logging.basicConfig(
    filename=AQUI / 'monitor.log',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s',
    encoding='utf-8'
)
log = logging.getLogger('optimus_mon')

HORARIOS_APRR = [6, 9, 11, 13, 15, 17]
JANELAS_TELE_MIN = [h * 60 - 30 for h in HORARIOS_APRR]  # 05:30, 08:30, 10:30, 12:30, 14:30, 16:30


def F(rotulo, caminho, esperado, contem='', so_hoje=True, regra='', conf_mensal=False,
      arquivos_obrigatorios=None, prefixos=None, caminho_fallback='', **extra):
    return dict(
        **extra,
        rotulo=rotulo,
        caminho=caminho,
        caminho_fallback=caminho_fallback,
        esperado=esperado,
        contem=contem,
        so_hoje=so_hoje,
        regra=regra,
        conf_mensal=conf_mensal,
        arquivos_obrigatorios=arquivos_obrigatorios or [],
        prefixos=prefixos or []
    )


B7 = '7 - BASES DE RELATORIOS\\'
MM = '\\{YYYY}\\{MM} - {MES}'

# ==============================================================================
# 2. CONFIGURAÇÃO PADRÃO (INCLUINDO PASTA LOCALHOST 30 MIN, VERCEL E WHATSAPP 6H)
# ==============================================================================
DEFAULT = {
    'versao_config': 4,
    'raiz': r'G:\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES',
    'raiz_rede_fallback': r'\\10.1.17.4\Usuarios\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES',
    'pasta_logs': r'1 - RELATORIOS\52 - LOGS PROJETOS',
    'pasta_relatorio_diario': r'G:\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES\7 - BASES DE RELATORIOS\RELATORIO PAINEL MONITORAMENTO DIARIO',
    'pasta_dashboard_acompanhamento': r'C:\Users\agenor.sales\Desktop\dashboard acompanhamento',
    'url_painel_vercel': 'https://acompanhameto-fluxo-diario-de-taref.vercel.app/',
    'intervalo_vercel_min': 30,  # Alimenta o painel online a cada 30 minutos
    'hora_relatorio_diario': 18,
    'intervalo_seg': 60,         # Atualiza a tela a cada 60 segundos
    'porta_servidor_web': 8080,
    'abrir_navegador_ao_iniciar': True,
    'dashboard_html': 'dashboard.html',
    'whatsapp': {
        'ativo': True,
        'porta_edge': 9222,
        'resumo_a_cada_min': 360,  # Envia no WhatsApp apenas a cada 6 horas (360 minutos)
        'horarios_resumo': ['00:00', '06:00', '12:00', '18:00', '23:59'],
        'alerta_imediato_erro': False,  # Desativado para respeitar rigorosamente a janela de 6 horas
        'perfil_edge': r'C:\edge_whatsapp_monitor',
        'edge_exe': '',
        'destinos': [
            {'nome': 'Agenor', 'numero': '558589768794'},
            {'nome': 'Miller', 'numero': '553199300997'}
        ]
    },
    'monitores': [
        {
            'nome': 'BI Alerta Vagas v2', 'categoria': 'bi', 'prio': 'P0', 'regra': 'aprr',
            'log': 'log_execucao_alerta_de_vagas_{YYYY}_{MM}.xlsx', 'min_linhas': 40,
            'pastas': [
                F(n, B7 + n + ' - DIARIA' + MM + '\\{DD}', 2, '', True, 'aprr', False)
                for n in ('T22APRR', 'T22APRR_PROC', 'T22APRR_VS', 'T22APRR_TRANSF')
            ]
        },
        {
            'nome': 'BI Alerta Vagas Tele Saúde', 'categoria': 'bi', 'prio': 'P0', 'regra': 'tele',
            'log': 'log_bi_telesaude_{YYYY}_{MM}.xlsx', 'min_linhas': 12,
            'pastas': [
                F(
                    rotulo='47 - BI_TELESSAUDE FATO',
                    caminho=r'\\10.1.17.4\Usuarios\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES\1 - RELATORIOS\47 - BI_TELESSAUDE\FATO',
                    caminho_fallback=r'1 - RELATORIOS\47 - BI_TELESSAUDE\FATO',
                    esperado=1,
                    contem='F_vagas_atual.parquet',
                    so_hoje=True,
                    regra='tele',
                    arquivos_obrigatorios=['F_vagas_atual.parquet']
                )
            ]
        },
        {
            'nome': 'BI Necessidade de Compra', 'categoria': 'bi', 'prio': 'P0', 'regra': '',
            'log': 'log_Necessidade_de_Compra_{YYYY}_{MM}.xlsx', 'min_linhas': 6,
            'pastas': [
                F(
                    rotulo='46 FATO',
                    caminho=r'1 - RELATORIOS\46 - BI NECESSIDADE COMPRA\FATO',
                    caminho_fallback=r'1 - RELATORIOS\46 - BI NECESSIDADE COMPRA FATO',
                    esperado=1,
                    so_hoje=True,
                    arquivos_obrigatorios=['BASE_FATO_NECESSIDADE_COMPRA.parquet']
                ),
                F(
                    rotulo='46 DIMENSOES',
                    caminho=r'1 - RELATORIOS\46 - BI NECESSIDADE COMPRA\DIMENSOES',
                    caminho_fallback=r'1 - RELATORIOS\46 - BI NECESSIDADE COMPRA DIMENSOES',
                    esperado=3,
                    so_hoje=True,
                    arquivos_obrigatorios=['DIM_CIDADE.parquet', 'DIM_ESPECIALIDADE.parquet', 'DIM_NECESSIDADE_COMPRA.parquet']
                )
            ]
        },
        {
            'nome': 'T0012 - Disponibilidade / Marcada', 'categoria': 'base', 'prio': 'P1', 'regra': '',
            'log': '', 'min_linhas': 0,
            'pastas': [
                F('T0012 DISPONIBILIDADE', B7 + r'T0012 - ULTIMA BASE DISPONIBILIDADE MENSAL\{YYYY}', 15, '', True, '', conf_mensal=True),
                F('T0012 MARCADA', B7 + r'T0012 - ULTIMA BASE MARCADA MENSAL\{YYYY}', 15, '', True, '', conf_mensal=True)
            ]
        },
        {
            'nome': 'T22J2 - Consultas Disponibilizadas', 'categoria': 'base', 'prio': 'P1', 'regra': '',
            'log': 'log_execucao_disponibilizadas_t22j2_{YYYY}_{MM}.xlsx', 'min_linhas': 5,
            'pastas': [
                F('T22J2 BASE MENSAL', B7 + r'T22J2 - BASE MENSAL\{YYYY}', 15, '', True, '', conf_mensal=True,
                  prefixos=['CCG', 'CLINIPAM', 'NDI', 'NDISPRJ', 'HAP'])
            ]
        },
        {
            'nome': 'T22J3 - Base diária', 'categoria': 'base', 'prio': 'P1', 'regra': '',
            'log': '', 'min_linhas': 0,
            'pastas': [
                F('T22J3', B7 + 'T22J3 - BASE DIARIA' + MM, 1, '{DD}{MM}{YYYY}', True, '', conf_mensal=True)
            ]
        },
        {
            'nome': 'T9016 - Base mensal', 'categoria': 'base', 'prio': 'P1', 'regra': '',
            'log': '', 'min_linhas': 0,
            'pastas': [
                F('T9016', B7 + 'T9016 - BASE MENSAL' + MM, 5, '{DD}{MM}{YYYY}', True, '', conf_mensal=True,
                  prefixos=['CCG', 'CLINIPAM', 'NDI', 'NDISPRJ'])
            ]
        },
        {
            'nome': 'T9033 - Atraso Médico', 'categoria': 'base', 'prio': 'P2', 'regra': '',
            'log': '', 'min_linhas': 0,
            'pastas': [
                F(
                    rotulo='ATRASO MEDICO',
                    caminho=B7 + 'BASE_PAINEL DOS MEDICOS\\BASE_ATRASO MEDICO' + MM,
                    caminho_fallback=B7 + 'BASE_PAINEL DOS MEDICOS BASE_ATRASO MEDICO' + MM,
                    esperado=1,
                    contem='',
                    so_hoje=True,
                    conf_mensal=True
                )
            ]
        },
        {
            'nome': 'T9033 - Horas Médicos', 'categoria': 'base', 'prio': 'P2', 'regra': '',
            'log': '', 'min_linhas': 0,
            'pastas': [
                F(
                    rotulo='HORAS MEDICOS',
                    caminho=B7 + 'BASE_PAINEL DOS MEDICOS\\BASE HORAS MEDICOS' + MM,
                    caminho_fallback=B7 + 'BASE PAINEL DOS MEDICOS BASE HORAS MEDICOS' + MM,
                    esperado=5,
                    contem='{DD}{MM}{YYYY}',
                    so_hoje=True,
                    conf_mensal=True,
                    prefixos=['CCG', 'CLINIPAM', 'NDI', 'NDISPRJ']
                )
            ]
        },
        {
            'nome': 'Estoque e Abastecimento', 'categoria': 'bi', 'prio': 'P1', 'regra': '',
            'log': 'log_execucao_estoque_abastecimento_{YYYY}_{MM}.xlsx', 'min_linhas': 4, 'pastas': []
        },
        {'nome': 'BI Painel dos Médicos', 'categoria': 'bi', 'prio': 'P1', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'BI Farol ICH (DGE)', 'categoria': 'bi', 'prio': 'P1', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'BI Relatório de Consultas Disponibilizadas', 'categoria': 'bi', 'prio': 'P1', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'BI Controle Agendas', 'categoria': 'bi', 'prio': 'P1', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'BI Captação de Agendas', 'categoria': 'bi', 'prio': 'P1', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'BI INDICADORES REDE CRED', 'categoria': 'bi', 'prio': 'P1', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'Logs Projetos (Automações)', 'categoria': 'bi', 'prio': 'P2', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'BI Relatório de Consultas em Transferência', 'categoria': 'bi', 'prio': 'P2', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'BI Relatório de Cirurgias', 'categoria': 'bi', 'prio': 'P2', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'BI Relatório Atraso dos Médicos', 'categoria': 'bi', 'prio': 'P2', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'BI Relatório de Acompanhamento de Mensageria', 'categoria': 'bi', 'prio': 'P2', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'BI Relatorio_SAC', 'categoria': 'bi', 'prio': 'P2', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []},
        {'nome': 'CAPTADOS FORA DE ALERTA', 'categoria': 'bi', 'prio': 'P2', 'regra': '', 'log': '', 'min_linhas': 0, 'pastas': []}
    ]
}


# ==============================================================================
# NOVAS REGRAS (documento MAPEAMENTO_DAS_PASTAS atualizado)
# ==============================================================================
def _mon(nome):
    return next(x for x in DEFAULT['monitores'] if x['nome'] == nome)


def _bi(nome, busca, prio, regra='diaria', **kw):
    return dict(nome=nome, busca=busca, prio=prio, regra=regra, **kw)


DEFAULT['versao_config'] = 5
DEFAULT['pasta_pbix'] = r'3 - ATUALIZAÇÃO PBIX\PBIX OFICIAL'
DEFAULT['hora_limite_manha'] = '11:00'          # painéis da manhã: atualizados até 11h
DEFAULT['bi_checar_fim_de_semana'] = False
DEFAULT['logs_ignorar'] = []
DEFAULT['equipe'] = ['Karine', 'Agenor', 'Miller']     # Daniel removido da equipe
DEFAULT['intervalo_pdf_horas'] = 2                      # PDF de saúde da equipe a cada 2 horas
DEFAULT['pastas_pdf_saude'] = [                         # cópias extras do PDF (a cópia local fica em .\pdf_saude_equipe)
    r'C:\Users\agenor.sales\Desktop\dashboard acompanhamento\PDF SAUDE EQUIPE',
    r'G:\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES\7 - BASES DE RELATORIOS\RELATORIO PAINEL MONITORAMENTO DIARIO\PDF SAUDE EQUIPE',
]                    # trechos de nome de logs que não devem ser auditados
_APRR_H = ['06:00', '09:00', '11:00', '13:00', '15:00', '17:00']
DEFAULT['bis_pbix'] = [
    _bi('Alerta de Vagas v2', ['alerta', 'vagas', 'v2'], 'P0', 'horarios', horarios=_APRR_H, tolerancia_min=150, excluir=['sub', 'tele'], resp='Karine'),
    _bi('Alerta Diário de Vagas_Sub', ['alerta', 'vagas', 'sub'], 'P0', 'horarios', horarios=_APRR_H, tolerancia_min=150, resp='Karine'),
    _bi('BI Alerta Vagas Tele Saúde', ['vagas', 'tele'], 'P0', 'horarios', horarios=_APRR_H, tolerancia_min=150, resp='Karine'),
    _bi('BI Necessidade de Compra', ['necessidade', 'compra'], 'P0', resp='Agenor'),
    _bi('Painel Médicos Hapclinicas V2', ['painel', 'medicos'], 'P1', resp='Karine'),
    _bi('Relatório Indicadores Rede Cred', ['indicadores', 'rede'], 'P1', resp='Karine'),
    _bi('Relatório de Consultas Disponibilizadas v2', ['consultas', 'disponibilizadas'], 'P1', resp='Karine'),
    _bi('Captação de Agendas', ['captacao', 'agendas'], 'P1', 'ate_hora', hora_limite='17:00', tolerancia_min=30, resp='Karine'),
    _bi('Controle de Agendas', ['controle', 'agendas'], 'P1', 'ate_hora', hora_limite='17:00', tolerancia_min=30, resp='Karine'),
    _bi('Farol de Índice Consultas por Hora - ICH', ['farol'], 'P1', 'ate_hora', hora_limite='14:00', tolerancia_min=30, resp='Karine'),
    _bi('Relatório de Estoque e Abastecimento v2', ['estoque', 'abastecimento'], 'P1', resp='Agenor'),
    _bi('Acompanhamento de Pagamentos Prestadores', ['pagamentos', 'prestadores'], 'P1', 'ultima', resp='Miller'),
    _bi('Escala Especial', ['escala', 'especial'], 'P2', resp='Miller'),
    _bi('Relatório Atraso dos Médicos', ['atraso', 'medicos'], 'P2', resp='Karine'),
    _bi('Relatório de Consultas em Transferência', ['transferencia'], 'P2', resp='Karine'),
    _bi('Relatório de Cancelamento v2', ['cancelamento'], 'P2', resp='Karine'),
    _bi('Relatório de Acompanhamento de Mensageria', ['mensageria'], 'P2', resp='Karine'),
    _bi('Monitoramento de Marcação', ['monitoramento', 'marcacao'], 'P2', 'ate_hora', hora_limite='06:00', tolerancia_min=60, resp='Karine'),
    _bi('Relatório de Cirurgias', ['cirurgias'], 'P2', resp='Karine'),
    _bi('Relatório SAC (mensal)', ['sac'], 'P2', 'mensal', excluir=['nip'], resp='Karine'),
    _bi('Médicos Novos (mensal)', ['medicos', 'novos'], 'P2', 'mensal', resp='Karine'),
    _bi('Relatório de SAC e NIP (desenvolvimento)', ['sac', 'nip'], 'P2', 'horarios', horarios=['07:30', '09:00', '13:00', '22:30'], tolerancia_min=90, resp=''),  # Daniel saiu da equipe: sem responsável
    _bi('Relatório de Tagueamento', ['tagueamento'], 'P2', 'horarios', horarios=['08:00', '11:30', '17:00'], tolerancia_min=90, resp='Miller'),
    _bi('Acompanhamento de Valores Pagos por Especialidade', ['valores', 'pagos', 'especialidade'], 'P2', 'ate_hora', hora_limite='05:00', tolerancia_min=60, resp='Agenor'),
    _bi('Forecast Monitoramento Limites e Mediana', ['forecast'], 'P2', 'ultima', resp='Miller e Agenor'),
    _bi('Relatório Check-up da Mulher (versão antiga)', ['check', 'mulher'], 'P2', 'ultima'),
    _bi('Censo_Medicos', [], 'P2', 'pausa'),
    _bi('Relatório Captação de Médicos', [], 'P2', 'pausa'),
    _bi('Relatórios Governança de Agendas', [], 'P2', 'pausa'),
    _bi('Relatório Pacotes', ['pacotes'], 'P2', 'ignorar'),
]
_RESP = {
    'BI Alerta Vagas v2': 'Karine', 'BI Alerta Vagas Tele Saúde': 'Karine', 'BI Necessidade de Compra': 'Agenor', 'Estoque e Abastecimento': 'Agenor',
    'T0012 - Disponibilidade / Marcada': 'Agenor', 'T22J2 - Consultas Disponibilizadas': 'Karine', 'T9033 - Atraso Médico': 'Karine',
    'T9033 - Horas Médicos': 'Karine', 'BI Painel dos Médicos': 'Karine', 'BI Farol ICH (DGE)': 'Karine', 'BI Relatório de Consultas Disponibilizadas': 'Karine',
    'BI Controle Agendas': 'Karine', 'BI Captação de Agendas': 'Karine', 'BI INDICADORES REDE CRED': 'Karine', 'BI Relatório de Consultas em Transferência': 'Karine',
    'BI Relatório de Cirurgias': 'Karine', 'BI Relatório Atraso dos Médicos': 'Karine', 'BI Relatório de Acompanhamento de Mensageria': 'Karine',
    'BI Relatorio_SAC': 'Karine', 'CAPTADOS FORA DE ALERTA': 'Karine',
}

# Bases: T0012 passa a 14 arquivos/dia e ganha o DB mensal (3 arquivos)
for _m in DEFAULT['monitores']:
    _m['responsavel'] = _RESP.get(_m['nome'], '')
for _p in _mon('T0012 - Disponibilidade / Marcada')['pastas']:
    _p['esperado'] = 14
DEFAULT['monitores'].append({
    'nome': 'DB T0012 (mensal)', 'categoria': 'base', 'prio': 'P1', 'regra': '', 'log': '', 'min_linhas': 0,
    'pastas': [F('DB_T0012 FLUXO PYTHON', B7 + r'DB_T0012\FLUXO PYTHON', 3, '', True, '', conf_mensal=True)]
})
# T9016: só baixa nos dias 1, 5 e 10; 5 arquivos no mês (um por empresa)
_p = _mon('T9016 - Base mensal')['pastas'][0]
_p.update(esperado=5, contem='', conf_mensal=False, dias_mes=[1, 5, 10],
          prefixos=['T9016_', 'NDISPRJ', 'NDIMINAS', 'CLINIPAM', 'CCG'])
# Indicadores Rede Cred: base atualizada até as 10h
DEFAULT['monitores'].append({
    'nome': 'Indicadores Rede Cred (bases)', 'categoria': 'base', 'prio': 'P1', 'regra': '', 'log': '', 'min_linhas': 0,
    'pastas': [
        F('CLIENTE_OCULTO HISTORICO', B7 + r'BI_INDICADORES_CRED\CLIENTE_OCULTO\HISTORICO', 1, '', True, '', conf_mensal=True, hora_limite='10:00'),
        F('CLIENTE_OCULTO MENSAL', B7 + r'BI_INDICADORES_CRED\CLIENTE_OCULTO', 1, '', True, '', conf_mensal=True, hora_limite='10:00'),
    ]
})
# Base DGE: 1 arquivo do mês atual em cada pasta
DEFAULT['monitores'].append({
    'nome': 'Base DGE (produtividade / pagamento)', 'categoria': 'base', 'prio': 'P2', 'regra': '', 'log': '', 'min_linhas': 0,
    'pastas': [
        F('48 - PRODUTIVIDADE V2 DGE', r'1 - RELATORIOS\48 - RELATORIO DE PRODUTIVIDADE V2\DGE', 1, '', False, '', so_mes=True),
        F('BASES DGE PRODUTIVIDADE', B7 + r'BASES DGE\PRODUTIVIDADE', 1, '', False, '', so_mes=True),
        F('BASES DGE TIPO PAGAMENTO', B7 + r'BASES DGE\TIPO PAGAMENTO', 1, '', False, '', so_mes=True),
    ]
})
_mon('DB T0012 (mensal)')['responsavel'] = 'Agenor'
_mon('Indicadores Rede Cred (bases)')['responsavel'] = 'Karine'
# Logs (Pasta 52) e janelas de execução dos BIs que antes eram só cartões pré-cadastrados
for _n, _log, _min, _ini, _fim in [
    ('BI Painel dos Médicos', 'log_execucao_painel_medicos_{YYYY}_{MM}.xlsx', 8, '07:10', '11:30'),
    ('BI Farol ICH (DGE)', 'log_execucao_farol_de_indice_consultas_por_hora_ich_{YYYY}_{MM}.xlsx', 4, '14:50', '15:10'),
    ('BI Relatório de Consultas Disponibilizadas', 'log_execucao_consultas_disponibilizadas_{YYYY}_{MM}.xlsx', 5, '02:00', '03:30'),
    ('BI Controle Agendas', 'log_controle_agendas_{YYYY}_{MM}.xlsx', 5, '16:50', '17:30'),
    ('BI Relatório de Consultas em Transferência', 'log_execucao_consultas_em_transferencia_{YYYY}_{MM}.xlsx', 4, '22:00', '22:30'),
    ('BI Relatório Atraso dos Médicos', 'log_Atraso_dos_Medicos_{YYYY}_{MM}.xlsx', 6, '07:30', '09:00'),
    ('BI Relatório de Acompanhamento de Mensageria', 'log_relatorio_mensageria_{YYYY}_{MM}.xlsx', 4, '22:00', '22:30'),
    ('CAPTADOS FORA DE ALERTA', 'log_Script_Tabela_Captados_Fora_Alerta_{YYYY}_{MM}.xlsx', 8, '01:00', '01:30'),
]:
    _mon(_n).update(log=_log, min_linhas=_min, hora_inicio=_ini, hora_fim=_fim)


# ==============================================================================
# 3. AUTOCONFIGURAÇÃO E RESOLUÇÃO DE CAMINHOS
# ==============================================================================
def tok(s, d=None):
    d = d or date.today()
    return (
        (s or '')
        .replace('{YYYY}', str(d.year))
        .replace('{MM}', f'{d.month:02d}')
        .replace('{DD}', f'{d.day:02d}')
        .replace('{MES}', MES[d.month - 1])
    )


def sincronizar_config_automatico():
    cfg_final = json.loads(json.dumps(DEFAULT))
    if CFG.exists():
        try:
            antigo = json.loads(CFG.read_text(encoding='utf-8'))
            if 'whatsapp' in antigo and 'destinos' in antigo['whatsapp']:
                cfg_final['whatsapp']['destinos'] = antigo['whatsapp']['destinos']
        except Exception:
            pass
    # Garante sempre 6 horas (360 min) no WhatsApp e 30 min para a Vercel
    cfg_final['whatsapp']['resumo_a_cada_min'] = 360
    cfg_final['intervalo_vercel_min'] = 30
    cfg_final['intervalo_pdf_horas'] = 2
    CFG.write_text(json.dumps(cfg_final, indent=2, ensure_ascii=False), encoding='utf-8')
    return cfg_final


def montar_caminho(raiz, rel, d=None):
    rel_resolved = tok(rel, d)
    if re.match(r'^([A-Za-z]:|\\\\)', rel_resolved):
        return Path(rel_resolved)
    return Path(raiz).joinpath(*[p for p in re.split(r'[\\/]', rel_resolved) if p])


def resolver_pasta_automatica(cfg, rel_principal, rel_fallback='', d=None):
    raizes = [
        cfg.get('raiz', DEFAULT['raiz']),
        cfg.get('raiz_rede_fallback', DEFAULT['raiz_rede_fallback']),
        r'G:\Credenciamento Medico\NUCLEO DE AGENDAS RAFAEL FERNANDES'
    ]
    candidatos = []
    for r in raizes:
        if rel_principal:
            candidatos.append(montar_caminho(r, rel_principal, d))
        if rel_fallback:
            candidatos.append(montar_caminho(r, rel_fallback, d))

    for c in candidatos:
        try:
            if c.exists():
                return c
        except Exception:
            continue
    return candidatos[0] if candidatos else Path('.')


def calcular_dias_uteis_mes(ate_data=None):
    hoje = ate_data or date.today()
    dias = []
    d = date(hoje.year, hoje.month, 1)
    while d <= hoje:
        if d.weekday() < 5:
            dias.append(d)
        d += timedelta(days=1)
    return dias


def hm(m):
    return '-' if m is None else f'{(m // 60) % 24:02d}:{m % 60:02d}'


# ==============================================================================
# 4. REGRAS DE HORÁRIO: APRR (+2 A CADA HORÁRIO) E TELE SAÚDE (30M ANTES)
# ==============================================================================
def calcular_esperado_aprr(agora=None):
    agora = agora or datetime.now()
    min_atual = agora.hour * 60 + agora.minute
    janelas_passadas = [h for h in HORARIOS_APRR if min_atual >= (h * 60 + 10)]
    qtd_esperada = len(janelas_passadas) * 2
    prox = next((h for h in HORARIOS_APRR if h * 60 > min_atual), None)
    return {
        'esperado': qtd_esperada,
        'janelas': len(janelas_passadas),
        'proximo_horario': f'{prox:02d}:00' if prox else 'Encerrado hoje (12/12)'
    }


def verificar_janela_telesaude(ult_mod_dt, agora=None):
    agora = agora or datetime.now()
    min_atual = agora.hour * 60 + agora.minute
    passadas = [m for m in JANELAS_TELE_MIN if min_atual >= m]
    prox = next((m for m in JANELAS_TELE_MIN if m > min_atual), None)
    prox_str = hm(prox) if prox else 'Amanhã às 05:30'

    if not passadas:
        return True, 'Aguardando 1ª janela às 05:30 (30m antes das 06h)'

    ult_janela = passadas[-1]
    ult_janela_str = hm(ult_janela)
    hora_aprr_ref = f'{(ult_janela + 30) // 60:02d}:00'

    if not ult_mod_dt:
        return False, f'Sem atualização na janela das {ult_janela_str} (30m antes das {hora_aprr_ref})'

    mod_hoje = (ult_mod_dt.date() == agora.date())
    mod_min = ult_mod_dt.hour * 60 + ult_mod_dt.minute
    em_dia = mod_hoje and (mod_min >= (ult_janela - 15))
    if em_dia:
        return True, f'Atualizado na janela das {ult_janela_str} (às {ult_mod_dt:%H:%M:%S}) | Próxima: {prox_str}'
    return False, f'ATRASADO: Esperada atualização às {ult_janela_str} (30m antes das {hora_aprr_ref}). Última: {ult_mod_dt:%d/%m %H:%M:%S}'


# ==============================================================================
# 4B. MONITORAMENTO DOS BIs (PBIX OFICIAL) + AUDITORIA DOS LOGS DA PASTA 52
# ==============================================================================
import calendar
import unicodedata

_CACHE_LOGS = {'t': 0.0, 'item': None, 'datas': {}}


def _norm(s):
    s = unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().lower()
    return ' ' + re.sub(r'[^a-z0-9]+', ' ', s).strip() + ' '


def _hora_dt(h, d):
    a, b = h.split(':')
    return datetime.combine(d, dtime(int(a), int(b)))


def _item_base(nome, prio, categoria):
    agora = datetime.now()
    return dict(
        nome=nome, categoria=categoria, prio=prio, regra='', status='ok', erros=[], avisos=[], infos=[], pastas=[],
        quando=agora.strftime('%d/%m/%Y %H:%M:%S'), quando_hms=agora.strftime('%H:%M:%S'),
        linhas_hoje=None, ok_hoje=None, err_hoje=None, media_min=None, media_arq_min=None, ultimo_fim=None, eventos_hoje=[]
    )


def _regra_pbix(b, mt, agora):
    """Retorna (status, motivo, esperado, atraso_min) conforme a regra do BI."""
    hoje = agora.date()
    regra = b.get('regra', 'diaria')
    if regra == 'mensal':
        ok = (mt.year, mt.month) == (hoje.year, hoje.month)
        return ('ok' if ok else 'err', f'Atualizado em {mt:%d/%m/%Y}' if ok else f'Sem atualização neste mês (última: {mt:%d/%m/%Y})', 'Atualização mensal', 0)
    if regra == 'ultima':
        dias = (agora - mt).days
        lim = int(b.get('limite_dias', 3))
        return ('ok' if dias <= lim else 'warn', f'Última atualização há {dias} dia(s)', f'Conferir última atualização (limite {lim} dias)', max(0, (dias - lim) * 1440))
    if hoje.weekday() >= 5 and not b.get('checar_fim_de_semana', False):
        return 'ok', 'Fim de semana: sem exigência diária', 'Dias úteis', 0
    if regra == 'horarios':
        tol = int(b.get('tolerancia_min', 90))
        sl = [_hora_dt(h, hoje) for h in b['horarios']]
        passadas = [s for s in sl if s + timedelta(minutes=tol) <= agora]
        prox = next((s for s in sl if s > agora), None)
        esp = 'Horários: ' + ', '.join(b['horarios']) + (f' | próximo {prox:%H:%M}' if prox else ' | encerrado hoje')
        if not passadas:
            return 'ok', 'Aguardando o 1º horário do dia', esp, 0
        ult = passadas[-1]
        if mt >= ult - timedelta(minutes=5):
            return 'ok', f'Atualizado após o horário das {ult:%H:%M} (às {mt:%H:%M:%S})', esp, 0
        atraso = int((agora - ult).total_seconds() // 60)
        return 'err', f'ATRASADO: esperada atualização do horário das {ult:%H:%M} (+{tol} min). Última: {mt:%d/%m %H:%M:%S}', esp, atraso
    prazo = _hora_dt(b['hora_limite'], hoje) + timedelta(minutes=int(b.get('tolerancia_min', 0)))
    esp = f"Atualizar até {b['hora_limite']}" + (f" (+{b['tolerancia_min']} min)" if b.get('tolerancia_min') else '')
    if mt.date() == hoje:
        return 'ok', f'Atualizado hoje às {mt:%H:%M:%S}', esp, 0
    if agora < prazo:
        return 'ok', f'Aguardando atualização (prazo {prazo:%H:%M})', esp, 0
    return 'err', f'ATRASADO: sem atualização hoje (prazo {prazo:%H:%M}). Última: {mt:%d/%m/%Y %H:%M:%S}', esp, int((agora - prazo).total_seconds() // 60)


def avaliar_pbix(cfg):
    agora = datetime.now()
    pasta = resolver_pasta_automatica(cfg, cfg.get('pasta_pbix', ''))
    try:
        arquivos = [(e, datetime.fromtimestamp(e.stat().st_mtime)) for e in pasta.iterdir()
                    if e.is_file() and e.suffix.lower() in ('.pbix', '.pbit')] if pasta.is_dir() else None
    except Exception:
        arquivos = None
    if arquivos is None:
        r = _item_base('PBIX · Pasta PBIX OFICIAL', 'P0', 'pbix')
        r['responsavel'] = 'Karine'
        r['status'] = 'err'
        r['erros'].append(f'Pasta dos BIs não encontrada ou ilegível: {pasta}')
        return [r]

    itens = []
    for b in cfg.get('bis_pbix', []):
        if b.get('regra') == 'ignorar':
            continue
        b = {'hora_limite': cfg.get('hora_limite_manha', '11:00'), 'checar_fim_de_semana': cfg.get('bi_checar_fim_de_semana', False), **b}
        r = _item_base('PBIX · ' + b['nome'], b['prio'], 'pbix')
        r['responsavel'] = b.get('resp', ''); r['tipo_proc'] = 'pbix'
        if b.get('regra') == 'pausa':
            r['status'] = 'nd'; r['pausa'] = True; r['infos'].append('Relatório em pausa (sem monitoramento)')
            r['pbix'] = dict(arquivo=None, ultima=None, ultima_iso=None, idade_h=None, esperado='Em pausa', atraso_min=0, regra='pausa')
            itens.append(r)
            continue
        ach = [(e, mt) for e, mt in arquivos
               if all(k in _norm(e.name) for k in b['busca']) and not any(k in _norm(e.name) for k in b.get('excluir', []))]
        if not ach:
            r['status'] = 'warn'
            r['avisos'].append(f"Arquivo .pbix não localizado (busca: {', '.join(b['busca'])}). Ajuste 'busca' em bis_pbix.")
            r['pbix'] = dict(arquivo=None, ultima=None, ultima_iso=None, idade_h=None, esperado='-', atraso_min=0, regra=b['regra'])
        else:
            e, mt = max(ach, key=lambda x: x[1])
            st, motivo, esp, atraso = _regra_pbix(b, mt, agora)
            if st == 'err' and b['prio'] == 'P2':
                st = 'warn'
            r['status'] = st
            {'err': r['erros'], 'warn': r['avisos']}.get(st, r['infos']).append(motivo)
            if len(ach) > 1:
                r['infos'].append(f'{len(ach)} arquivos casam com a busca; usando o mais recente ({e.name})')
            r['pbix'] = dict(arquivo=e.name, ultima=mt.strftime('%d/%m/%Y %H:%M:%S'), ultima_iso=mt.isoformat(timespec='seconds'),
                             idade_h=round((agora - mt).total_seconds() / 3600, 1), esperado=esp, atraso_min=atraso, regra=b['regra'])
        itens.append(r)
    return itens


def _dia_util_anterior(d):
    d = d - timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


def auditar_logs_52(cfg, forcar=False):
    """Regra Pasta 52: log do mês atual atualizado diariamente; logs de meses anteriores com a última data do mês;
    família de log cuja última data está 3+ meses atrás da referência do mês = DESATUALIZADO."""
    if not forcar and _CACHE_LOGS['item'] and time.time() - _CACHE_LOGS['t'] < 600:
        item = json.loads(json.dumps(_CACHE_LOGS['item']))
        agora = datetime.now()
        item['quando'], item['quando_hms'] = agora.strftime('%d/%m/%Y %H:%M:%S'), agora.strftime('%H:%M:%S')
        return item

    hoje = date.today()
    r = _item_base('Auditoria de Logs (Pasta 52)', 'P1', 'logs')
    pasta = resolver_pasta_automatica(cfg, cfg['pasta_logs'])
    if not pasta.is_dir():
        r['status'] = 'err'
        r['erros'].append(f'Pasta 52 não encontrada: {pasta}')
        return r

    mes_hoje = hoje.year * 12 + hoje.month
    ignorar = [_norm(x) for x in cfg.get('logs_ignorar', [])]
    fams = {}
    for arq in pasta.glob('*.xlsx'):
        m = re.search(r'_(\d{4})_(\d{2})\.xlsx$', arq.name, re.I)
        if not m or arq.name.startswith('~$'):
            continue
        ref = (int(m[1]), int(m[2]))
        if mes_hoje - (ref[0] * 12 + ref[1]) > 4 or any(i.strip() in _norm(arq.name) for i in ignorar):
            continue
        mt = arq.stat().st_mtime
        chave = (arq.name, mt)
        if chave not in _CACHE_LOGS['datas']:
            try:
                ds = [x['d'] for x in ler_log(arq) if x['d']]
                _CACHE_LOGS['datas'][chave] = max(ds) if ds else None
            except Exception:
                _CACHE_LOGS['datas'][chave] = None
        fams.setdefault(arq.name[:m.start()], []).append(dict(arquivo=arq.name, ref=ref, ult=_CACHE_LOGS['datas'][chave]))

    det, util_ant = [], _dia_util_anterior(hoje)
    for fam, arqs in fams.items():
        arqs.sort(key=lambda x: x['ref'])
        for i, a in enumerate(arqs):
            ult, ref = a['ult'], a['ref']
            ult_s = ult.strftime('%d/%m/%Y') if ult else '-'
            ultimo_da_familia = (i == len(arqs) - 1)
            if ref == (hoje.year, hoje.month):
                if ult == hoje or (hoje.weekday() >= 5 and ult and ult >= util_ant):
                    st, msg = 'ok', 'Atualizado hoje'
                elif ult and ult >= util_ant:
                    st, msg = 'warn', f'Última data {ult_s} (dia útil anterior); ainda sem execução hoje'
                else:
                    st, msg = 'err', f'DESATUALIZADO: última data {ult_s}, esperado hoje'
            else:
                fim = date(ref[0], ref[1], calendar.monthrange(*ref)[1])
                meses_atras = (mes_hoje - (ult.year * 12 + ult.month)) if ult else 99
                if ultimo_da_familia and meses_atras >= 3:
                    st, msg = 'err', f'DESATUALIZADO: última data {ult_s} (3+ meses atrás) e sem log do mês atual'
                elif ultimo_da_familia:
                    st, msg = 'warn', f'Sem log do mês atual (último: {ref[1]:02d}/{ref[0]}, até {ult_s})'
                elif ult and ult >= fim - timedelta(days=int(cfg.get('log_tolerancia_fim_mes_dias', 3))):
                    st, msg = 'ok', f'Fechou o mês em {ult_s}'
                else:
                    st, msg = 'warn', f'Não chegou ao fim do mês (última data {ult_s})'
            det.append(dict(arquivo=a['arquivo'], ref=f'{ref[1]:02d}/{ref[0]}', ultima=ult_s, status=st, motivo=msg))

    ordem = {'err': 0, 'warn': 1, 'ok': 2}
    det.sort(key=lambda x: (ordem[x['status']], x['arquivo']))
    r['logs_detalhe'] = det[:80]
    errs = [d for d in det if d['status'] == 'err']
    warns = [d for d in det if d['status'] == 'warn']
    if errs:
        r['status'] = 'err'
        r['erros'] += [f"{d['arquivo']}: {d['motivo']}" for d in errs[:8]]
        if len(errs) > 8:
            r['erros'].append(f'... e mais {len(errs) - 8} log(s) desatualizado(s)')
    elif warns:
        r['status'] = 'warn'
    if warns:
        r['avisos'].append(f'{len(warns)} log(s) com alerta (ver lista abaixo)')
    r['infos'].append(f'{len(det)} logs analisados na Pasta 52')
    _CACHE_LOGS.update(t=time.time(), item=json.loads(json.dumps(r)))
    return r


# ==============================================================================
# 5. LEITURA DE LOGS (.XLSX) E AVALIAÇÃO DOS MONITORES
# ==============================================================================
def ler_log(arq):
    from openpyxl import load_workbook
    wb = load_workbook(arq, read_only=True, data_only=True)
    saida = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            cel = [c for c in row if c is not None and str(c).strip() != '']
            txt = [str(c).strip() for c in cel]
            si = next((i for i, t in enumerate(txt) if re.match(r'(?i)^(SUCESSO|ERRO|EXECUTANDO|RODANDO|EM ANDAMENTO)', t)), None)
            if si is None:
                continue
            d, tempos, tempos_hms = None, [], []
            for c in cel:
                if isinstance(c, datetime):
                    d = d or c.date()
                    if c.hour or c.minute or c.second:
                        tempos.append(c.hour * 60 + c.minute)
                        tempos_hms.append(c.strftime('%H:%M:%S'))
                elif isinstance(c, dtime):
                    tempos.append(c.hour * 60 + c.minute)
                    tempos_hms.append(c.strftime('%H:%M:%S'))
                elif isinstance(c, date):
                    d = d or c
                else:
                    t = str(c).strip()
                    if not re.fullmatch(r'[\d/\-: ]+', t):
                        continue
                    m = re.search(r'(\d{2})/(\d{2})/(\d{4})', t)
                    if m and d is None:
                        d = date(int(m[3]), int(m[2]), int(m[1]))
                    m = re.search(r'(\d{4})-(\d{2})-(\d{2})', t)
                    if m and d is None:
                        d = date(int(m[1]), int(m[2]), int(m[3]))
                    for a, b, s in re.findall(r'\b(\d{1,2}):(\d{2})(?::(\d{2}))?\b', t):
                        tempos.append(int(a) * 60 + int(b))
                        tempos_hms.append(f'{int(a):02d}:{int(b):02d}:{int(s or 0):02d}')

            ini, fim = (tempos + [None, None])[:2]
            ini_hms, fim_hms = (tempos_hms + [hm(ini) + ':00' if ini else '-', hm(fim) + ':00' if fim else '-'])[:2]
            dur = None
            if ini is not None and fim is not None:
                dur = fim - ini + (1440 if fim < ini else 0)
            script = next((t for t in txt if re.search(r'\.(py|yxmd)$', t, re.I)), txt[2] if len(txt) > 2 else txt[0])
            saida.append(dict(
                d=d, ini=ini, fim=fim, ini_hms=ini_hms, fim_hms=fim_hms,
                dur=dur, st=txt[si].upper()[:3], script=script, msg=' | '.join(txt[si + 1:])
            ))
    wb.close()
    return saida


def avaliar(m, cfg):
    agora = datetime.now()
    hoje = agora.date()
    regra = m.get('regra', '')
    dias_uteis = calcular_dias_uteis_mes(hoje)
    qtd_dias_uteis = len(dias_uteis)

    r = dict(
        nome=m['nome'],
        categoria=m.get('categoria', 'bi'),
        prio=m['prio'],
        regra=regra,
        status='ok',
        erros=[],
        avisos=[],
        infos=[],
        pastas=[],
        quando=agora.strftime('%d/%m/%Y %H:%M:%S'),
        quando_hms=agora.strftime('%H:%M:%S'),
        linhas_hoje=None,
        ok_hoje=None,
        err_hoje=None,
        media_min=None,
        media_arq_min=None,
        ultimo_fim=None,
        eventos_hoje=[]
    )

    ult_mod_log = None
    if m.get('log'):
        pasta_52 = resolver_pasta_automatica(cfg, cfg['pasta_logs'])
        arq = pasta_52 / tok(m['log'], hoje)
        if not arq.exists():
            r['status'] = 'err'
            r['erros'].append(f'Log do mês não encontrado: {arq.name}')
        else:
            try:
                ult_mod_log = datetime.fromtimestamp(arq.stat().st_mtime)
                linhas = ler_log(arq)
                td = [x for x in linhas if x['d'] == hoje]
                r.update(
                    linhas_hoje=len(td),
                    ok_hoje=sum(x['st'] == 'SUC' for x in td),
                    err_hoje=sum(x['st'] == 'ERR' for x in td),
                    eventos_hoje=[{'h': hm(x['ini']) if x['ini'] is not None else '00:00', 'hms': x['ini_hms'], 'st': x['st']} for x in td]
                )
                dd = [x['dur'] for x in linhas if x['st'] == 'SUC' and x['dur'] is not None]
                r['media_min'] = round(sum(dd) / len(dd), 2) if dd else None
                r['ultimo_fim'] = td[-1]['fim_hms'] if td else '-'

                ultimo = {x['script']: x for x in td}
                if not td:
                    if regra == 'aprr' and calcular_esperado_aprr(agora)['esperado'] == 0:
                        r['infos'].append('Aguardando 1º horário da APRR (06:00)')
                    elif regra == 'tele' and agora.hour * 60 + agora.minute < JANELAS_TELE_MIN[0]:
                        r['infos'].append('Aguardando 1ª janela do Tele Saúde (05:30)')
                    elif m.get('hora_inicio') and agora.strftime('%H:%M') < m['hora_inicio']:
                        r['infos'].append(f"Aguardando execução prevista para {m['hora_inicio']}")
                    else:
                        r['status'] = 'warn'
                        r['avisos'].append('Nenhuma execução registrada hoje no log')
                else:
                    rod = [x for x in ultimo.values() if x['st'] in ('EXE', 'ROD', 'EM ')]
                    if rod:
                        r['status'] = 'run'
                        r['infos'].append('Rodando: ' + ', '.join(x['script'] for x in rod))
                    for x in ultimo.values():
                        if x['st'] == 'ERR':
                            r['status'] = 'err'
                            r['erros'].append(f"{x['script']} ({x['ini_hms']}): {x['msg'] or 'erro sem mensagem'}")
                    reex = [x for x in td if x['st'] == 'ERR' and ultimo[x['script']] is not x]
                    if reex:
                        r['avisos'].append(f'{len(reex)} erro(s) hoje já reexecutado(s) com sucesso')
                    if r['status'] == 'ok' and len(td) < m.get('min_linhas', 0) and regra not in ('aprr', 'tele') and agora.strftime('%H:%M') >= m.get('hora_fim', '00:00'):
                        r['status'] = 'warn'
                        r['avisos'].append(f"Log com {len(td)} linhas hoje; esperado {m['min_linhas']}+")
            except Exception as e:
                r['status'] = 'err'
                r['erros'].append(f'Falha ao ler o log: {e}')

    aprr_info = calcular_esperado_aprr(agora) if regra == 'aprr' else None
    if aprr_info:
        r['infos'].append(
            f"Regra APRR (06h,09h,11h,13h,15h,17h): {aprr_info['janelas']}/6 janelas -> Esperado agora: {aprr_info['esperado']} arqs/pasta"
        )

    ult_mod_pasta = None
    tempos_por_arq = []
    total_arquivos_hoje = 0

    for f in m.get('pastas', []):
        eh_aprr = (regra == 'aprr' or f.get('regra') == 'aprr')
        esperado_dia = aprr_info['esperado'] if eh_aprr else f['esperado']
        if f.get('dias_mes'):
            esperado_dia = 1 if hoje.day in f['dias_mes'] else 0
        if f.get('hora_limite') and agora.strftime('%H:%M') < f['hora_limite']:
            esperado_dia = 0
        conf_mensal = f.get('conf_mensal', False)
        esperado_mes = (f['esperado'] * qtd_dias_uteis) if conf_mensal else None

        o = dict(
            rotulo=f['rotulo'],
            esperado=esperado_dia,
            n=0,
            conf_mensal=conf_mensal,
            n_mes=0,
            esperado_mes=esperado_mes,
            dias_uteis_ok=0,
            total_dias_uteis=qtd_dias_uteis,
            erro=None,
            tempo_medio_arq=None
        )

        try:
            pasta = resolver_pasta_automatica(cfg, f['caminho'], f.get('caminho_fallback', ''), hoje)
            if not pasta.is_dir():
                if eh_aprr and esperado_dia == 0:
                    o['n'] = 0
                else:
                    raise FileNotFoundError(f'Pasta não encontrada: {pasta}')
            else:
                chave_hoje = tok(f.get('contem', ''), hoje).lower()
                n_hoje = 0
                n_mes = 0
                mtimes_hoje = []
                datas_com_arquivo_no_mes = set()

                for e in pasta.iterdir():
                    if not e.is_file():
                        continue
                    mtime_dt = datetime.fromtimestamp(e.stat().st_mtime)
                    nome_low = e.name.lower()

                    if mtime_dt.year == hoje.year and mtime_dt.month == hoje.month:
                        n_mes += 1
                        datas_com_arquivo_no_mes.add(mtime_dt.date())
                    for du in dias_uteis:
                        if f"{du.day:02d}{du.month:02d}{du.year}" in nome_low:
                            datas_com_arquivo_no_mes.add(du)

                    mod_hoje = (mtime_dt.date() == hoje)
                    tem_chave = (chave_hoje in nome_low) if chave_hoje else False

                    if chave_hoje:
                        if tem_chave or mod_hoje:
                            n_hoje += 1
                            mtimes_hoje.append(mtime_dt)
                    elif not f.get('so_hoje', True) or mod_hoje:
                        n_hoje += 1
                        mtimes_hoje.append(mtime_dt)

                if f.get('so_mes'):
                    n_hoje = n_mes
                o['n'] = n_hoje
                o['n_mes'] = n_mes
                total_arquivos_hoje += n_hoje

                dias_uteis_anteriores = [d for d in dias_uteis if d < hoje]
                dias_faltantes = [d.strftime('%d/%m') for d in dias_uteis_anteriores if d not in datas_com_arquivo_no_mes]
                o['dias_uteis_ok'] = qtd_dias_uteis - len(dias_faltantes)

                if conf_mensal and dias_faltantes:
                    if r['status'] == 'ok':
                        r['status'] = 'warn'
                    r['avisos'].append(
                        f"{f['rotulo']}: Falta arquivo em {len(dias_faltantes)} dia(s) útil(eis) do mês ({', '.join(dias_faltantes[:5])})"
                    )

                if mtimes_hoje:
                    mtimes_hoje.sort()
                    if ult_mod_pasta is None or mtimes_hoje[-1] > ult_mod_pasta:
                        ult_mod_pasta = mtimes_hoje[-1]
                    if r['media_min'] and n_hoje > 0:
                        o['tempo_medio_arq'] = round(r['media_min'] / n_hoje, 2)
                    elif len(mtimes_hoje) > 1:
                        intervalo_min = (mtimes_hoje[-1] - mtimes_hoje[0]).total_seconds() / 60.0
                        o['tempo_medio_arq'] = round(max(0.1, intervalo_min / len(mtimes_hoje)), 2)
                    else:
                        o['tempo_medio_arq'] = round(r['media_min'] or 1.0, 2)
                    tempos_por_arq.append(o['tempo_medio_arq'])

                if f.get('dias_mes') and f.get('prefixos') and (hoje.day > min(f['dias_mes']) or agora.hour >= 12):
                    tok_mes = tok('{MM}{YYYY}', hoje)
                    nomes_mes = [e.name.lower() for e in pasta.iterdir() if e.is_file() and tok_mes in e.name]
                    faltam = [p for p in f['prefixos'] if not any(p.lower() in nm for nm in nomes_mes)]
                    if faltam:
                        r['status'] = 'err'
                        r['erros'].append(f"NÃO BAIXOU: {f['rotulo']} sem arquivo de {', '.join(faltam)} (baixa nos dias {f['dias_mes']})")
                if n_hoje < esperado_dia:
                    r['status'] = 'err'
                    r['erros'].append(f"NÃO BAIXOU: {f['rotulo']} tem {n_hoje} de {esperado_dia} arquivos hoje")
        except Exception as e:
            o['erro'] = str(e)
            r['status'] = 'err'
            r['erros'].append(f"{f['rotulo']}: {e}")
        r['pastas'].append(o)

    if r['media_min'] and total_arquivos_hoje > 0:
        r['media_arq_min'] = round(r['media_min'] / total_arquivos_hoje, 2)
    elif tempos_por_arq:
        r['media_arq_min'] = round(sum(tempos_por_arq) / len(tempos_por_arq), 2)

    if regra == 'tele':
        datas_ref = [d for d in (ult_mod_log, ult_mod_pasta) if d is not None]
        maior_data = max(datas_ref) if datas_ref else None
        ok_tele, msg_tele = verificar_janela_telesaude(maior_data, agora)
        if not ok_tele:
            r['status'] = 'err'
            r['erros'].append(msg_tele)
        else:
            r['infos'].append(msg_tele)

    if not m.get('log') and not m.get('pastas'):
        r['status'] = 'nd'
        r['infos'].append('Card pré-cadastrado no painel OptimusBI')
    return r


# ==============================================================================
# 6. GERAÇÃO DO .XLSX E AUTOMACAO DE UPLOAD NA VERCEL A CADA 30 MINUTOS
# ==============================================================================
def gerar_relatorio_xlsx(cfg, res, snaps_hoje, hist_diario, salvar_em_pasta_especifica=None):
    """Gera o arquivo .xlsx na pasta solicitada (ou tanto em Desktop\\dashboard acompanhamento quanto em G:\\...)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    agora = datetime.now()
    nome_arq = f"Relatorio_OptimusBI_{agora:%Y-%m-%d_%Hh%Mm%Ss}.xlsx"

    wb = Workbook()
    header_fill = PatternFill(start_color="0339A6", end_color="0339A6", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    ws1 = wb.active
    ws1.title = "Resumo_Diario"
    cab1 = [
        "Data/Hora (HH:MM:SS)", "Categoria", "Processo", "Prioridade", "Status",
        "Execuções Hoje", "Sucessos Hoje", "Erros Hoje", "Último Fim (HH:MM:SS)",
        "Tempo Médio Processo (min)", "Tempo Médio / Arquivo (min)", "Meta Diária / Mensal", "Alertas / Erros"
    ]
    ws1.append(cab1)
    for col_num in range(1, len(cab1) + 1):
        c = ws1.cell(row=1, column=col_num)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center")

    for r in res:
        metas_txt = " | ".join(
            f"{p['rotulo']}: Dia {p['n']}/{p['esperado']}" + (f" (Mês {p['n_mes']}/{p['esperado_mes']})" if p.get('conf_mensal') else "")
            for p in r['pastas']
        )
        ws1.append([
            r['quando'],
            {'bi': 'Painel BI', 'pbix': 'BI (.pbix)', 'logs': 'Auditoria de Logs'}.get(r['categoria'], 'Base de Dados'),
            r['nome'],
            r['prio'],
            r['status'].upper(),
            r['linhas_hoje'] if r['linhas_hoje'] is not None else 0,
            r['ok_hoje'] if r['ok_hoje'] is not None else 0,
            r['err_hoje'] if r['err_hoje'] is not None else 0,
            r['ultimo_fim'] or '-',
            r['media_min'] if r['media_min'] is not None else '-',
            r['media_arq_min'] if r['media_arq_min'] is not None else '-',
            metas_txt or '-',
            " | ".join(r['erros'] + r['avisos'])
        ])

    ws2 = wb.create_sheet("Estado_Por_Hora")
    cab2 = ["Hora do Registro", "Horário Exato (HH:MM:SS)", "Processo", "Categoria", "Status na Hora", "Qtd Erros", "Mensagens de Falha"]
    ws2.append(cab2)
    for col_num in range(1, len(cab2) + 1):
        c = ws2.cell(row=1, column=col_num)
        c.fill = header_fill
        c.font = header_font

    for hora_chave in sorted(snaps_hoje.keys()):
        snap = snaps_hoje[hora_chave]
        for item in snap.get('itens', []):
            ws2.append([
                hora_chave,
                snap.get('salvo_as', '-'),
                item['nome'],
                item['categoria'],
                item['status'].upper(),
                item.get('err_hoje') or 0,
                " | ".join(item.get('erros', []))
            ])

    ws3 = wb.create_sheet("Evolucao_Por_Dia")
    cab3 = ["Data", "Hora Última Leitura", "Processos OK", "Processos em Alerta", "Processos com Erro"]
    ws3.append(cab3)
    for col_num in range(1, len(cab3) + 1):
        c = ws3.cell(row=1, column=col_num)
        c.fill = header_fill
        c.font = header_font

    for h in hist_diario:
        ws3.append([h.get('dia'), h.get('h'), h.get('ok'), h.get('warn'), h.get('err')])

    # 1. Salva sempre na pasta C:\Users\agenor.sales\Desktop\dashboard acompanhamento
    pasta_local_acomp = Path(cfg.get('pasta_dashboard_acompanhamento', DEFAULT['pasta_dashboard_acompanhamento']))
    try:
        pasta_local_acomp.mkdir(parents=True, exist_ok=True)
    except Exception:
        pasta_local_acomp = AQUI

    arq_local_datado = pasta_local_acomp / nome_arq
    arq_local_fixo = pasta_local_acomp / "monitoramento_atual.xlsx"
    wb.save(arq_local_datado)
    wb.save(arq_local_fixo)
    log.info('Planilha salva em %s', arq_local_datado)

    # 2. Se for o relatório diário ou solicitado, salva também na rede G:\...
    if salvar_em_pasta_especifica:
        try:
            p_rede = Path(salvar_em_pasta_especifica)
            p_rede.mkdir(parents=True, exist_ok=True)
            wb.save(p_rede / nome_arq)
        except Exception as e:
            log.warning('Não foi possível salvar cópia na rede %s: %s', salvar_em_pasta_especifica, e)

    return arq_local_datado


def enviar_xlsx_para_painel_vercel(cfg, caminho_xlsx):
    """
    Abre o Microsoft Edge no link https://acompanhameto-fluxo-diario-de-taref.vercel.app/,
    clica na aba Monitoramento e inclui o arquivo .xlsx gerado.
    """
    from selenium import webdriver
    from selenium.webdriver.edge.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    url_vercel = cfg.get('url_painel_vercel', DEFAULT['url_painel_vercel'])
    porta_vercel = 9225  # Porta separada para não conflitar com o WhatsApp (9222)
    proc = None
    drv = None

    try:
        log.info('Iniciando upload do arquivo %s para o painel Vercel (%s)...', caminho_xlsx.name, url_vercel)
        encerrar_edge_porta(porta_vercel)

        w = cfg['whatsapp']
        cands = [
            w.get('edge_exe', ''),
            r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
            r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
            os.path.expandvars(r'%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe')
        ]
        exe = next((c for c in cands if c and os.path.exists(c)), None)
        if not exe:
            log.error('msedge.exe não encontrado para subir planilha na Vercel.')
            return False

        perfil_vercel = r'C:\edge_vercel_monitor'
        proc = subprocess.Popen([
            exe,
            f'--remote-debugging-port={porta_vercel}',
            f'--user-data-dir={perfil_vercel}',
            '--no-first-run',
            url_vercel
        ])

        for _ in range(35):
            if porta_aberta(porta_vercel):
                time.sleep(3)
                break
            time.sleep(1)

        op = Options()
        op.debugger_address = f"127.0.0.1:{porta_vercel}"
        drv = webdriver.Edge(options=op)
        drv.get(url_vercel)
        time.sleep(5)

        # 1. Procura e clica na aba "Monitoramento"
        xpaths_aba = [
            "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'monitoramento')]",
            "//button[contains(., 'Monitoramento')]",
            "//a[contains(., 'Monitoramento')]"
        ]
        for xp in xpaths_aba:
            elems = drv.find_elements(By.XPATH, xp)
            for el in elems:
                if el.is_displayed():
                    try:
                        el.click()
                        time.sleep(2)
                        break
                    except Exception:
                        pass

        # 2. Localiza o <input type="file"> na aba Monitoramento e envia o caminho do .xlsx
        inputs_file = drv.find_elements(By.XPATH, "//input[@type='file']")
        if not inputs_file:
            # Aguarda até 15s caso a aba ainda esteja renderizando
            inp = WebDriverWait(drv, 15).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
            )
        else:
            inp = inputs_file[0]

        # Torna o input visível caso esteja com classe 'hidden' no React/Tailwind
        drv.execute_script("arguments[0].style.display = 'block'; arguments[0].style.visibility = 'visible';", inp)
        time.sleep(1)
        inp.send_keys(str(caminho_xlsx.resolve()))
        time.sleep(6)  # Aguarda o processamento/sincronização do arquivo na aplicação

        # Se houver botão de confirmar/salvar/enviar após selecionar o arquivo, clica nele
        botoes_confirmar = drv.find_elements(
            By.XPATH,
            "//button[contains(., 'Salvar') or contains(., 'Enviar') or contains(., 'Importar') or contains(., 'Sincronizar')]"
        )
        for btn in botoes_confirmar:
            if btn.is_displayed():
                try:
                    btn.click()
                    time.sleep(3)
                except Exception:
                    pass

        log.info('Arquivo %s incluído com sucesso na aba Monitoramento da Vercel!', caminho_xlsx.name)
        return True

    except Exception as e:
        log.error('Falha ao incluir arquivo na Vercel: %s', str(e)[:250])
        return False
    finally:
        if proc:
            try:
                proc.terminate()
            except Exception:
                pass
        encerrar_edge_porta(porta_vercel)


def rotina_vercel_30_minutos(cfg, res, snaps_hoje, hist, forcar=False):
    """Executa a cada 30 minutos: gera o .xlsx na pasta do Desktop e inclui na aba Monitoramento da Vercel."""
    est = json.loads(ESTADO.read_text(encoding='utf-8')) if ESTADO.exists() else {}
    agora = datetime.now()
    ult_str = est.get('ultimo_upload_vercel')
    ult_dt = datetime.fromisoformat(ult_str) if ult_str else None
    int_min = int(cfg.get('intervalo_vercel_min', 30))

    if forcar or ult_dt is None or (agora - ult_dt) >= timedelta(minutes=int_min):
        arq_xlsx = gerar_relatorio_xlsx(cfg, res, snaps_hoje, hist)
        enviar_xlsx_para_painel_vercel(cfg, arq_xlsx)
        est['ultimo_upload_vercel'] = agora.isoformat()
        ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')


# ==============================================================================
# 7. TEMPLATE DO DASHBOARD LOCAL (HTML)
# ==============================================================================
VERSAO_PAINEL = 'v6.0 · Visão Geral | BIs | Logs | Detalhamento'

HTML = r'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Monitoramento OptimusBI</title>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap">
  <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
  <style>
    :root {
      --blue-inst: #0339A6;
      --blue-dark: #122A44;
      --gold: #F2B705;
      --orange: #F24405;
      --red: #F21D2F;
      --bg: #F2F2F2;
      --ink: #1F2937;
      --card: #FFFFFF;
      --slate-100: #F1F5F9;
      --slate-200: #E2E8F0;
      --slate-400: #94A3B8;
      --slate-600: #475569;
      --amber-50: #FFFBEB;
      --amber-800: #92400E;
      --ok: #10B981;
    }
    * { box-sizing: border-box; }
    body { margin: 0; background: var(--bg); color: var(--ink); font: 14px/1.45 'IBM Plex Sans', system-ui, sans-serif; }
    header { background: var(--blue-inst); color: var(--card); padding: 12px 24px 0; box-shadow: 0 2px 8px rgba(0,0,0,0.12); position: sticky; top: 0; z-index: 50; }
    .header-top { display: flex; flex-wrap: wrap; gap: 14px; align-items: center; justify-content: space-between; padding-bottom: 12px; }
    .brand { display: flex; align-items: center; gap: 12px; }
    .logo-box { width: 46px; height: 46px; background: var(--card); border-radius: 8px; border: 2px solid var(--gold); display: flex; align-items: center; justify-content: center; overflow: hidden; cursor: pointer; }
    .logo-box img { max-width: 100%; max-height: 100%; object-fit: contain; }
    h1 { font-size: 20px; margin: 0; font-weight: 700; color: var(--card); }
    .sub { color: var(--slate-200); font-size: 12.5px; display: flex; align-items: center; gap: 8px; margin-top: 2px; flex-wrap: wrap; }
    .badge { background: var(--blue-dark); color: var(--gold); padding: 2px 8px; border-radius: 4px; font-family: monospace; font-weight: 700; font-size: 13px; }
    .bar { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
    button, select { font: inherit; font-size: 13px; font-weight: 600; border: 1px solid var(--slate-200); background: var(--card); color: var(--ink); padding: 7px 14px; border-radius: 6px; cursor: pointer; transition: 0.15s; }
    button:hover { background: var(--blue-dark); color: var(--card); border-color: var(--blue-dark); }
    button.p { background: var(--gold); color: var(--blue-dark); border-color: var(--gold); }
    button.p:hover { background: var(--blue-dark); color: var(--gold); }
    button.g { background: var(--blue-dark); color: var(--card); border-color: var(--blue-dark); }
    .nav-tabs { display: flex; gap: 4px; border-top: 1px solid rgba(255,255,255,0.15); padding-top: 8px; overflow-x: auto; }
    .nav-tab { background: transparent; color: var(--slate-200); border: none; border-bottom: 3px solid transparent; border-radius: 6px 6px 0 0; padding: 8px 18px; font-weight: 600; cursor: pointer; }
    .nav-tab:hover { background: var(--blue-dark); color: var(--card); }
    .nav-tab.active { background: var(--bg); color: var(--blue-inst); border-bottom-color: var(--gold); }
    main { padding: 20px 24px; max-width: 1440px; margin: auto; }
    .filter-bar { background: var(--card); border: 1px solid var(--slate-200); border-left: 5px solid var(--blue-inst); border-radius: 8px; padding: 10px 16px; margin-bottom: 16px; display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 10px; }
    .sum { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 16px; }
    .k { background: var(--card); border: 1px solid var(--slate-200); border-left: 5px solid var(--c); border-radius: 8px; padding: 12px 14px; }
    .k b { font-size: 26px; display: block; line-height: 1.1; color: var(--ink); }
    .k span { color: var(--slate-600); font-size: 12px; }
    .charts { display: grid; grid-template-columns: repeat(auto-fit, minmax(400px, 1fr)); gap: 14px; margin-bottom: 18px; }
    .box { background: var(--card); border: 1px solid var(--slate-200); border-radius: 8px; padding: 14px; }
    .box h3 { margin: 0 0 10px; font-size: 13.5px; color: var(--blue-inst); font-weight: 700; }
    .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(420px, 1fr)); gap: 14px; }
    .m { background: var(--card); border: 1px solid var(--slate-200); border-top: 4px solid var(--c); border-radius: 8px; padding: 15px; }
    .m:hover { border-color: var(--blue-dark); }
    .m h2 { font-size: 15.5px; margin: 0; color: var(--blue-inst); display: flex; justify-content: space-between; gap: 8px; }
    .chip { font-size: 11px; font-weight: 700; color: #fff; background: var(--c); border-radius: 99px; padding: 2px 10px; white-space: nowrap; }
    .meta { display: flex; gap: 6px; margin: 4px 0 10px; font-size: 11px; font-weight: 700; }
    .p-tag { padding: 2px 8px; border-radius: 4px; color: #fff; }
    .p-P0 { background: var(--red); } .p-P1 { background: var(--orange); } .p-P2 { background: var(--gold); color: var(--ink); } .p-P3 { background: var(--blue-inst); }
    .c-tag { background: var(--slate-100); color: var(--slate-600); padding: 2px 8px; border-radius: 4px; }
    .row { display: flex; justify-content: space-between; gap: 10px; padding: 4px 0; border-bottom: 1px dashed var(--slate-200); font-size: 13px; }
    .row:last-of-type { border: 0; } .row span:first-child { color: var(--slate-600); }
    .msg { margin-top: 8px; padding: 8px 10px; border-radius: 6px; background: #FEE2E2; color: var(--red); border-left: 3px solid var(--red); font-size: 12.5px; font-weight: 500; }
    .msg.w { background: var(--amber-50); color: var(--amber-800); border-left-color: var(--gold); }
    .msg.i { background: var(--slate-100); color: var(--blue-inst); border-left-color: var(--blue-inst); }
    .ok { --c: var(--ok); } .err { --c: var(--red); } .warn { --c: var(--gold); } .run { --c: var(--blue-inst); } .nd { --c: var(--slate-400); }
    dialog { border: 0; border-radius: 10px; width: min(880px, 95vw); padding: 0; } dialog::backdrop { background: rgba(18,42,68,0.65); }
    .dh { padding: 14px 18px; background: var(--blue-inst); color: #fff; display: flex; justify-content: space-between; align-items: center; }
    .db { padding: 16px 18px; background: #fff; }
    .hidden { display: none !important; }
    .alert-badge { position: fixed; right: 18px; bottom: 18px; z-index: 50; background: var(--red); color: #fff; border: 0; border-radius: 99px; padding: 10px 18px; font: 700 13px 'IBM Plex Sans', sans-serif; cursor: pointer; box-shadow: 0 4px 14px rgba(242,29,47,.45); animation: pulso 1.6s infinite; }
    @keyframes pulso { 0%,100% { transform: scale(1); } 50% { transform: scale(1.06); } }
    #dAlert { border-top: 6px solid var(--red); width: min(720px, 95vw); }
    #dAlert .al { border-left: 4px solid var(--c); background: var(--slate-100); border-radius: 6px; padding: 8px 12px; margin-bottom: 8px; font-size: 13px; }
    #dAlert .al b { color: var(--blue-dark); } #dAlert .al small { display: block; color: var(--slate-600); margin-top: 2px; }
    .lt { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 11.5px; }
    .lt td, .lt th { border-bottom: 1px solid var(--slate-200); padding: 3px 4px; text-align: left; }
    .lt .s-err { color: var(--red); font-weight: 700; } .lt .s-warn { color: var(--amber-800); font-weight: 600; } .lt .s-ok { color: var(--blue-inst); }
    .box.big { padding: 18px; } .box.big h3 { font-size: 15px; }
    .cw { position: relative; height: 380px; } .charts2 { display: grid; grid-template-columns: repeat(auto-fit, minmax(560px, 1fr)); gap: 16px; margin-bottom: 18px; }
    .grid1 { display: block; } .sum .k b { font-size: 32px; }
    .pb { margin-bottom: 12px; } .pb-h { display: flex; justify-content: space-between; font-size: 13.5px; margin-bottom: 4px; } .pb-h span { color: var(--slate-600); }
    .pb-t { background: var(--slate-200); border-radius: 99px; height: 20px; overflow: hidden; }
    .pb-f { height: 100%; width: 0; border-radius: 99px; transition: width 1.4s cubic-bezier(.22,.9,.3,1); }
    .pb0 .pb-t { height: 28px; } .pb0 .pb-h { font-size: 15px; }
    .chips { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 14px; }
    .chip-f { border: 1px solid var(--blue-inst); background: #fff; color: var(--blue-inst); border-radius: 99px; padding: 6px 14px; font: 600 12.5px 'IBM Plex Sans', sans-serif; cursor: pointer; }
    .chip-f span { background: var(--slate-100); color: var(--slate-600); border-radius: 99px; padding: 0 7px; margin-left: 4px; }
    .chip-f.on { background: var(--blue-inst); color: #fff; } .chip-f.on span { background: rgba(255,255,255,.25); color: #fff; }
    #toasts { position: fixed; right: 14px; top: 150px; z-index: 60; display: flex; flex-direction: column; gap: 8px; width: 290px; pointer-events: none; }
    .toast { pointer-events: auto; position: relative; background: #fff; border-left: 5px solid var(--c); border-radius: 8px; padding: 8px 26px 8px 10px; box-shadow: 0 4px 14px rgba(18,42,68,.28); font-size: 12px; animation: entra .35s ease-out; }
    .toast .t1 { font-weight: 700; color: var(--blue-dark); } .toast .t2 { color: var(--slate-600); margin: 2px 0; } .toast .t3 { color: var(--ink); }
    .toast .x { position: absolute; right: 4px; top: 2px; border: 0; background: none; font-size: 18px; cursor: pointer; color: var(--slate-600); padding: 0 6px; }
    @keyframes entra { from { transform: translateX(120%); opacity: 0; } to { transform: none; opacity: 1; } }
    @media(max-width:600px){ #toasts { width: 240px; } .charts2 { grid-template-columns: 1fr; } }
    @media(max-width:600px){ .grid, .charts { grid-template-columns: 1fr; } }
  </style>
</head>
<body>
<header>
  <div class="header-top">
    <div class="brand">
      <div class="logo-box" id="logoBox" title="Clique para selecionar a logo OptimusBI">
        <img id="logoImg" src="logo.png" alt="OptimusBI" onerror="this.style.display='none';this.nextElementSibling.style.display='block'">
        <span style="display:none;color:var(--blue-inst);font-weight:800;font-size:10px;text-align:center">OPTIMUS<br>BI</span>
      </div>
      <input type="file" id="logoFile" accept="image/*" class="hidden">
      <div>
        <h1>Monitoramento OptimusBI</h1>
        <div class="sub">
          <span id="last"></span>
          <span>| Relógio: <b class="badge" id="liveClock">--:--:--</b></span>
          <span>| Atualização: <b class="badge" id="cd">60s</b></span>
          <span>| Painel: <b class="badge">__VERSAO__</b></span>
        </div>
      </div>
    </div>
    <div class="bar">
      <button class="p" id="bRefresh">🔄 Atualizar Agora</button>
      <button class="g" id="bVercel">☁️ Enviar p/ Vercel Agora</button>
      <button class="g" id="bRel">📊 Salvar Relatório (.xlsx)</button>
      <button class="g" id="bPdf">📄 PDF Saúde da Equipe</button>
      <button class="g" id="bWa">💬 Resumo no WhatsApp</button>
      <button id="bBell" title="Reexibir alertas dispensados">🔔 <span id="bellN">0</span></button>
    </div>
  </div>
  <nav class="nav-tabs" id="tabs">
    <button class="nav-tab active" data-tab="geral">Visão Geral</button>
    <button class="nav-tab" data-tab="pbix">📊 Monitoramento de BIs</button>
    <button class="nav-tab" data-tab="logs">🗂️ Logs (Pasta 52)</button>
    <button class="nav-tab" data-tab="det">🧩 Detalhamento dos Processos</button>
  </nav>
</header>

<main>
  <div class="filter-bar">
    <div>
      <strong>⏱️ Filtro de Estado por Hora: </strong>
      <select id="hourSel"><option value="now">Tempo Real (Última Leitura)</option></select>
      <span id="snapInfo" style="font-size:12px;color:var(--slate-600);margin-left:8px"></span>
    </div>
    <div style="font-size:12px;color:var(--slate-600)">
      <strong>Vercel (.xlsx):</strong> a cada 30 min | <strong>WhatsApp:</strong> a cada 6 horas | <strong>PDF Saúde da Equipe:</strong> a cada 2 horas | <strong>Relatório Final:</strong> 18:00
    </div>
  </div>

  <section id="s-geral">
    <div class="sum" id="sum"></div>
    <div class="box big" style="margin-bottom:16px"><h3>Monitoramento dos BIs — atualizações por hora (área empilhada: cada cor = um BI)</h3><div id="biNote" style="font-size:12px;color:var(--slate-600);margin:-4px 0 8px"></div><div class="cw" id="cwBis"><canvas id="cBis"></canvas></div></div>
    <div class="box big" style="margin-bottom:16px">
      <h3>Evolução da equipe OptimusBI — % de itens em dia por responsável</h3>
      <div id="team"></div>
    </div>
    <div class="box big" style="margin-bottom:16px"><h3>Evolução da equipe por hora (hoje) — % de feito (itens em dia), equipe e cada responsável</h3><div class="cw"><canvas id="cEqHora"></canvas></div></div>
    <div class="charts2">
      <div class="box big"><h3>Falhas × tempo de resolução por hora (hoje) — linhas</h3><div class="cw"><canvas id="cHoraE"></canvas></div></div>
      <div class="box big"><h3>Evolução por dia — % em dia da equipe e de cada responsável (linhas)</h3><div class="cw"><canvas id="cDiaE"></canvas></div></div>
    </div>
  </section>

  <section id="s-pbix" class="hidden">
    <div class="chips" id="chipsBi"></div>
    <div class="grid" id="gridBi"></div>
  </section>

  <section id="s-logs" class="hidden">
    <div class="sum" id="sumLogs"></div>
    <div class="grid1" id="gridLogs"></div>
  </section>

  <section id="s-det" class="hidden">
    <div class="chips" id="chipsDet"></div>
    <div class="grid" id="gridDet"></div>
  </section>
</main>

<div id="toasts"></div>

<dialog id="dWa">
  <div class="dh"><b>Resumo para WhatsApp (A cada 6 horas)</b><button id="bWc">Fechar</button></div>
  <div class="db">
    <p style="font-size:12px;color:var(--slate-600);margin-top:0">O Python envia automaticamente a cada 6 horas (1 mensagem por contato, fechando e reiniciando o navegador).</p>
    <textarea id="waTxt" rows="11" style="width:100%;font:inherit;padding:8px"></textarea>
    <div class="bar" id="waBtns" style="margin-top:10px"></div>
  </div>
</dialog>

<script>
const D = __DATA__;
const $ = s => document.querySelector(s);
const $$ = s => document.querySelectorAll(s);
const p2 = n => String(n).padStart(2, '0');
const fmtHMS = (d = new Date()) => `${p2(d.getHours())}:${p2(d.getMinutes())}:${p2(d.getSeconds())}`;
const L = { ok: 'OK', err: 'ERRO', warn: 'ALERTA', run: 'RODANDO', nd: 'PRÉ-CADASTRADO' };
const CAT = { bi: 'Fluxo de execução', base: 'Download de base', pbix: 'BI (.pbix oficial)', logs: 'Auditoria de Logs' };
const COR = { ok: '#0339A6', warn: '#F2B705', err: '#F21D2F', nd: '#94A3B8' };
const PESSOA = { Karine: '#F24405', Agenor: '#F2B705', Miller: '#122A44' };
const EQUIPE = D.equipe || ['Karine', 'Agenor', 'Miller'];
const resp = r => r.responsavel || 'Equipe OptimusBI';
const donos = r => (r.responsavel || '').split(/\s+e\s+|,/).map(x => x.trim()).filter(x => x && EQUIPE.includes(x));
let tab = 'geral', fDet = 'todos', fBi = 'todos', charts = {}, secLeft = 60;

const savedLogo = localStorage.getItem('optimusLogo');
if (savedLogo) { $('#logoImg').src = savedLogo; $('#logoImg').style.display = 'block'; }
$('#logoBox').onclick = () => $('#logoFile').click();
$('#logoFile').onchange = e => {
  const f = e.target.files[0]; if (!f) return;
  const rd = new FileReader();
  rd.onload = () => { localStorage.setItem('optimusLogo', rd.result); location.reload(); };
  rd.readAsDataURL(f);
};
const snaps = D.snaps || {};
Object.keys(snaps).sort().forEach(h => {
  const s = snaps[h];
  const errs = (s.itens || []).filter(x => x.status === 'err').length;
  $('#hourSel').insertAdjacentHTML('beforeend', `<option value="${h}">Estado às ${h} (Registrado às ${s.salvo_as} - ${errs} erro(s))</option>`);
});
$('#hourSel').onchange = render;
$$('#tabs .nav-tab').forEach(b => b.onclick = () => {
  $$('#tabs .nav-tab').forEach(x => x.classList.remove('active'));
  b.classList.add('active'); tab = b.dataset.tab; render();
});
$('#bRefresh').onclick = async () => {
  $('#bRefresh').disabled = true; $('#bRefresh').textContent = '⏳ Atualizando...';
  try { await fetch('/api/atualizar'); } catch (e) {}
  location.reload();
};
$('#bVercel').onclick = async () => {
  $('#bVercel').disabled = true; $('#bVercel').textContent = '⏳ Enviando p/ Vercel...';
  try { const j = await (await fetch('/api/vercel')).json(); alert('Planilha gerada e enviada para a aba Monitoramento da Vercel!\nArquivo: ' + j.arquivo); }
  catch (e) { alert('O upload automático ocorre a cada 30 minutos pelo Python.'); }
  $('#bVercel').disabled = false; $('#bVercel').textContent = '☁️ Enviar p/ Vercel Agora';
};
$('#bPdf').onclick = async () => {
  $('#bPdf').disabled = true; $('#bPdf').textContent = '⏳ Gerando PDF...';
  try { const j = await (await fetch('/api/pdf')).json(); alert(j.ok ? 'PDF de saúde da equipe gerado em:\n' + j.arquivo : 'Não foi possível gerar o PDF. Veja o monitor.log (confira se reportlab e matplotlib estão instalados).'); }
  catch (e) { alert('O PDF é gerado automaticamente a cada 2 horas pelo Python.'); }
  $('#bPdf').disabled = false; $('#bPdf').textContent = '📄 PDF Saúde da Equipe';
};
$('#bRel').onclick = async () => {
  try { const j = await (await fetch('/api/relatorio')).json(); alert('Relatório .xlsx salvo em:\n' + j.arquivo); }
  catch (e) { alert('Relatório salvo automaticamente na pasta configurada!'); }
};

// ---------- cartão de detalhamento ----------
function logsTable(ls, tudo) {
  const sh = tudo ? ls : ls.filter(l => l.status !== 'ok').slice(0, 25);
  if (!sh.length) return '<div class="msg i">Todos os logs analisados estão em dia.</div>';
  return `<table class="lt"><tr><th>Log</th><th>Ref.</th><th>Última data</th><th>Situação</th></tr>` +
    sh.map(l => `<tr><td>${l.arquivo}</td><td>${l.ref}</td><td>${l.ultima}</td><td class="s-${l.status}">${l.motivo}</td></tr>`).join('') + '</table>';
}
function card(r) {
  const w = [];
  w.push(['👤 Responsável', `<b>${resp(r)}</b>`]);
  if (r.linhas_hoje != null) {
    w.push(['Execuções hoje', `${r.linhas_hoje} (${r.ok_hoje} ok / ${r.err_hoje} erro)`], ['Último fim (HH:MM:SS)', r.ultimo_fim || '-'],
      ['Tempo médio processo', r.media_min == null ? '-' : r.media_min.toFixed(2) + ' min']);
  }
  (r.pastas || []).forEach(f => {
    const tArq = f.tempo_medio_arq != null ? ` (${f.tempo_medio_arq} min/arq)` : '';
    w.push([f.rotulo + ' (Meta Diária)', f.erro ? `<b style="color:var(--red)">${f.erro}</b>` : `<b>${f.n}/${f.esperado} hoje</b>${tArq} ` + (f.n >= f.esperado ? '✔' : '✖')]);
    if (f.conf_mensal) w.push([f.rotulo + ' (Meta Mensal Útil)', `${f.n_mes}/${f.esperado_mes} no mês (${f.dias_uteis_ok}/${f.total_dias_uteis} dias úteis) ` + (f.dias_uteis_ok >= f.total_dias_uteis ? '✔' : '⚠️')]);
  });
  if (r.pbix) {
    w.push(['Arquivo .pbix', r.pbix.arquivo || '<b>não localizado</b>'], ['Última atualização', `<b>${r.pbix.ultima || '-'}</b>`], ['Regra esperada', r.pbix.esperado || '-']);
    if (r.pbix.arquivo) w.push(['Tempo desde a atualização', r.pbix.idade_h + ' h'], ['Atraso sobre o esperado', r.pbix.atraso_min ? `<b style="color:var(--red)">${r.pbix.atraso_min} min</b>` : '0 min']);
  }
  w.push(['Verificado em (Hora:Min:Seg)', `<b>${r.quando}</b>`]);
  return `<div class="m ${r.status}">
    <h2><span>${r.nome}</span><span class="chip">${r.pausa ? 'EM PAUSA' : L[r.status]}</span></h2>
    <div class="meta"><span class="p-tag p-${r.prio}">Prioridade ${r.prio}</span><span class="c-tag">${CAT[r.categoria] || 'Outros'}</span></div>
    ${w.map(x => `<div class="row"><span>${x[0]}</span><span>${x[1]}</span></div>`).join('')}
    ${(r.erros || []).map(x => `<div class="msg">${x}</div>`).join('')}
    ${(r.avisos || []).map(x => `<div class="msg w">${x}</div>`).join('')}
    ${(r.infos || []).map(x => `<div class="msg i">${x}</div>`).join('')}
    ${r.logs_detalhe ? logsTable(r.logs_detalhe, false) : ''}
  </div>`;
}
const ordena = a => [...a].sort((x, y) => x.prio.localeCompare(y.prio) || x.nome.localeCompare(y.nome));
function chips(el, opcoes, atual, set) {
  $(el).innerHTML = opcoes.map(([k, t, n]) => `<button class="chip-f ${k === atual ? 'on' : ''}" data-k="${k}">${t} <span>${n}</span></button>`).join('');
  $$(el + ' .chip-f').forEach(b => b.onclick = () => { set(b.dataset.k); render(); });
}
const mk = (id, cfg) => { charts[id]?.destroy(); charts[id] = new Chart($('#' + id), cfg); };

// ---------- render ----------
function render() {
  const hVal = $('#hourSel').value;
  let list = D.r;
  if (hVal !== 'now' && snaps[hVal]) {
    list = snaps[hVal].itens;
    $('#snapInfo').innerHTML = `<b style="color:var(--orange)">Exibindo fotografia das ${hVal} (salva às ${snaps[hVal].salvo_as})</b>`;
  } else $('#snapInfo').textContent = 'Pastas lidas automaticamente pelo Python.';
  $$('main > section').forEach(s => s.classList.toggle('hidden', s.id !== 's-' + tab));
  const ativos = list.filter(r => r.status !== 'nd');
  if (tab === 'geral') renderGeral(list, ativos, hVal);
  if (tab === 'pbix') renderBis(list);
  if (tab === 'logs') renderLogs(list);
  if (tab === 'det') renderDet(list);
  renderToasts(hVal);
}

function renderGeral(list, ativos, hVal) {
  const c = { ok: 0, warn: 0, err: 0 };
  ativos.forEach(r => { if (c[r.status] != null) c[r.status]++; });
  const bis = ativos.filter(r => r.categoria === 'pbix'), bOk = bis.filter(r => r.status === 'ok').length;
  const saude = ativos.length ? Math.round(100 * c.ok / ativos.length) : 0;
  $('#sum').innerHTML = [
    ['Itens monitorados', ativos.length, 'var(--blue-inst)'], ['Em dia (OK)', c.ok, 'var(--blue-inst)'],
    ['Alertas (Média)', c.warn, 'var(--gold)'], ['Crítico / Erro', c.err, 'var(--red)'],
    ['BIs em dia', `${bOk}/${bis.length}`, 'var(--blue-dark)'], ['Saúde geral', saude + '%', saude >= 90 ? 'var(--blue-inst)' : saude >= 70 ? 'var(--gold)' : 'var(--red)']
  ].map(([a, b, col]) => `<div class="k" style="--c:${col}"><b>${b}</b><span>${a}</span></div>`).join('');

  // BIs: área empilhada — horas x atualizações de BI (uma série por BI)
  const snapOn = hVal !== 'now' && snaps[hVal];
  const hLim = snapOn ? parseInt(hVal) : 23;
  const hBi = h => +h.t.slice(11, 13);
  const ev = (D.bi_ev || []).filter(e => e.t.slice(0, 10) === D.hoje && hBi(e) <= hLim);
  const maxH = Math.max(1, snapOn ? hLim : new Date().getHours(), ...ev.map(hBi));
  const horasBi = Array.from({ length: maxH + 1 }, (_, i) => i);
  const nomesBi = [...new Set(ev.map(e => e.bi))].sort();
  const dsBi = nomesBi.map((n, i) => {
    const hue = Math.round(360 * i / Math.max(1, nomesBi.length)), dono = ev.find(e => e.bi === n).resp || 'Equipe';
    return { label: `${n} · ${dono}`, data: horasBi.map(h => ev.filter(e => e.bi === n && hBi(e) === h).length), fill: true, stack: 'bi', tension: .3, pointRadius: 2, borderWidth: 1.5,
      borderColor: `hsl(${hue},65%,42%)`, backgroundColor: `hsla(${hue},65%,50%,.72)` };
  });
  $('#cwBis').style.height = '440px';
  $('#biNote').textContent = ev.length ? `${ev.length} atualização(ões) de BI registrada(s) hoje, em ${nomesBi.length} BI(s).` : 'Nenhuma atualização de BI registrada hoje até o momento.';
  mk('cBis', {
    type: 'line',
    data: { labels: horasBi.map(h => p2(h) + 'h'), datasets: dsBi },
    options: { maintainAspectRatio: false, animation: false, interaction: { mode: 'index', intersect: false },
      plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } }, tooltip: { itemSort: (a, b) => b.raw - a.raw, filter: i => i.raw > 0 } },
      scales: { x: { title: { display: true, text: 'hora do dia' } }, y: { stacked: true, beginAtZero: true, ticks: { precision: 0 }, title: { display: true, text: 'atualizações de BI na hora' } } } }
  });

  // equipe: barras crescentes
  const eq = {};
  const add = (n, r) => { const e = eq[n] || (eq[n] = [0, 0]); e[1]++; if (r.status === 'ok') e[0]++; };
  ativos.forEach(r => donos(r).forEach(n => add(n, r)));
  const linhas = [['Equipe OptimusBI (geral)', ativos.filter(r => r.status === 'ok').length, ativos.length], ...Object.entries(eq).sort().map(([n, v]) => [n, v[0], v[1]])];
  $('#team').innerHTML = linhas.map(([n, ok, t], i) => {
    const pct = t ? Math.round(100 * ok / t) : 0, cor = pct >= 90 ? 'var(--blue-inst)' : pct >= 70 ? 'var(--gold)' : 'var(--red)';
    return `<div class="pb ${i === 0 ? 'pb0' : ''}"><div class="pb-h"><b>${i ? '👤 ' : '👥 '}${n}</b><span>${ok} de ${t} em dia · <b>${pct}%</b></span></div><div class="pb-t"><div class="pb-f" data-w="${pct}" style="background:${cor}"></div></div></div>`;
  }).join('');
  requestAnimationFrame(() => setTimeout(() => $$('.pb-f').forEach(el => el.style.width = el.dataset.w + '%'), 60));

  const hoje = D.hoje, H = D.hist || [];
  const pct = x => { const t = x.ok + x.warn + x.err; return t ? Math.round(100 * x.ok / t) : 0; };
  const porHora = {}; H.filter(x => x.dia === hoje).forEach(x => porHora[x.h.slice(0, 2)] = x);
  const resHora = {}; (D.inc || []).filter(i => i.dia === hoje).forEach(i => (resHora[i.h] = resHora[i.h] || []).push(i.min));
  const hk = [...new Set([...Object.keys(porHora), ...Object.keys(resHora)])].sort();
  mk('cHoraE', {
    type: 'line',
    data: { labels: hk.map(h => h + 'h'), datasets: [
      { label: 'Falhas (itens em erro)', data: hk.map(h => porHora[h] ? porHora[h].err : null), borderColor: '#F21D2F', backgroundColor: '#F21D2F', tension: .3, yAxisID: 'y', spanGaps: true, pointRadius: 3 },
      { label: 'Tempo de resolução (min, média)', data: hk.map(h => resHora[h] ? Math.round(10 * resHora[h].reduce((a, b) => a + b, 0) / resHora[h].length) / 10 : null), borderColor: '#0339A6', backgroundColor: '#0339A6', tension: .3, yAxisID: 'y1', spanGaps: true, pointRadius: 3 }] },
    options: { maintainAspectRatio: false, animation: false, interaction: { mode: 'index', intersect: false }, scales: { y: { beginAtZero: true, ticks: { precision: 0 }, title: { display: true, text: 'falhas' } }, y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'minutos para resolver' } } } }
  });
  // equipe por hora: % de feito
  const hkE = Object.keys(porHora).sort();
  mk('cEqHora', {
    type: 'line',
    data: { labels: hkE.map(h => h + 'h'), datasets: [
      { label: 'Equipe (geral)', data: hkE.map(h => pct(porHora[h])), borderColor: '#0339A6', backgroundColor: '#0339A6', borderWidth: 4, tension: .3, pointRadius: 3 },
      ...EQUIPE.map(n => ({ label: n, data: hkE.map(h => { const e = (porHora[h].eq || {})[n]; return e && e[1] ? Math.round(100 * e[0] / e[1]) : null; }), borderColor: PESSOA[n] || '#94A3B8', backgroundColor: PESSOA[n] || '#94A3B8', tension: .3, spanGaps: true, pointRadius: 3 }))] },
    options: { maintainAspectRatio: false, animation: false, interaction: { mode: 'index', intersect: false }, scales: { y: { min: 0, max: 100, title: { display: true, text: '% feito (itens em dia)' } }, x: { title: { display: true, text: 'hora do dia' } } } }
  });
  const porDia = {}; H.forEach(x => porDia[x.dia] = x);
  const dk = Object.keys(porDia).sort().slice(-20);
  const nomes = EQUIPE.filter(n => dk.some(d => (porDia[d].eq || {})[n]));
  mk('cDiaE', {
    type: 'line',
    data: { labels: dk.map(d => d.slice(8) + '/' + d.slice(5, 7)), datasets: [
      { label: 'Equipe (geral)', data: dk.map(d => pct(porDia[d])), borderColor: '#0339A6', backgroundColor: '#0339A6', borderWidth: 4, tension: .3 },
      ...nomes.map(n => ({ label: n, data: dk.map(d => { const e = (porDia[d].eq || {})[n]; return e && e[1] ? Math.round(100 * e[0] / e[1]) : null; }), borderColor: PESSOA[n] || '#94A3B8', backgroundColor: PESSOA[n] || '#94A3B8', tension: .3, borderDash: [5, 4], spanGaps: true }))] },
    options: { maintainAspectRatio: false, animation: false, scales: { y: { min: 0, max: 100, title: { display: true, text: '% em dia' } } } }
  });
}

function renderBis(list) {
  const bis = list.filter(r => r.categoria === 'pbix');
  const pessoas = [...new Set(bis.flatMap(donos))].sort();
  chips('#chipsBi', [['todos', 'Todos', bis.length], ...pessoas.map(p => [p, '👤 ' + p, bis.filter(r => donos(r).includes(p)).length]),
    ['falha', '🚨 Com falha/alerta', bis.filter(r => r.status === 'err' || r.status === 'warn').length], ['pausa', '⏸️ Em pausa', bis.filter(r => r.pausa).length]], fBi, k => fBi = k);
  const f = bis.filter(r => fBi === 'todos' ? true : fBi === 'pausa' ? r.pausa : fBi === 'falha' ? (r.status === 'err' || r.status === 'warn') : donos(r).includes(fBi));
  $('#gridBi').innerHTML = ordena(f).map(card).join('') || '<div class="msg i">Nenhum BI neste filtro.</div>';
}

function renderLogs(list) {
  const it = list.find(r => r.categoria === 'logs');
  if (!it) { $('#sumLogs').innerHTML = ''; $('#gridLogs').innerHTML = '<div class="msg i">Auditoria de logs ainda não executada.</div>'; return; }
  const ls = it.logs_detalhe || [], n = s => ls.filter(l => l.status === s).length;
  $('#sumLogs').innerHTML = [['Logs analisados', ls.length, 'var(--blue-inst)'], ['Em dia', n('ok'), 'var(--blue-inst)'], ['Com alerta', n('warn'), 'var(--gold)'], ['Desatualizados', n('err'), 'var(--red)']]
    .map(([a, b, col]) => `<div class="k" style="--c:${col}"><b>${b}</b><span>${a}</span></div>`).join('');
  $('#gridLogs').innerHTML = `<div class="box"><h3>Regra: log do mês atual atualizado diariamente · logs anteriores fechando o mês · 3+ meses sem atualizar = desatualizado</h3>${(it.erros || []).map(x => `<div class="msg">${x}</div>`).join('')}${logsTable(ls, true)}<div class="row" style="margin-top:8px"><span>Verificado em</span><span><b>${it.quando}</b></span></div></div>`;
}

function renderDet(list) {
  const base = list.filter(r => r.categoria !== 'logs');
  const tipo = r => r.categoria === 'pbix' ? 'pbix' : (r.tipo_proc || 'outros');
  const cn = k => base.filter(r => tipo(r) === k).length;
  chips('#chipsDet', [['todos', 'Todos', base.length], ['fluxo', '⚙️ Fluxo de execução', cn('fluxo')], ['download', '⬇️ Download de bases', cn('download')], ['pbix', '📊 BIs (.pbix)', cn('pbix')], ['outros', 'Outros', cn('outros')]], fDet, k => fDet = k);
  $('#gridDet').innerHTML = ordena(base.filter(r => fDet === 'todos' || tipo(r) === fDet)).map(card).join('') || '<div class="msg i">Nada neste filtro.</div>';
}

// ---------- pop-ups pequenos na lateral ----------
function renderToasts(hVal) {
  const al = ordena(D.r.filter(r => r.status === 'err'));
  const dis = JSON.parse(localStorage.getItem('optimusDismiss') || '{}'), seen = JSON.parse(sessionStorage.getItem('optimusSeen') || '{}'), now = Date.now();
  const key = r => r.nome + '|' + (r.erros[0] || '');
  const vis = al.filter(r => !(dis[key(r)] && now - dis[key(r)] < 30 * 60000));
  $('#bellN').textContent = al.length;
  if (hVal !== 'now') { $('#toasts').innerHTML = ''; return; }
  const cor = r => r.prio === 'P0' ? 'var(--red)' : r.prio === 'P1' ? 'var(--orange)' : 'var(--gold)';
  $('#toasts').innerHTML = vis.slice(0, 5).map(r => {
    const e = (r.erros[0] || 'Falha detectada'); 
    return `<div class="toast" style="--c:${cor(r)}" data-k="${encodeURIComponent(key(r))}"><button class="x" title="Dispensar por 30 min">×</button>
      <div class="t1">🚨 [${r.prio}] ${r.nome.replace('PBIX · ', '')}</div><div class="t2">${e.length > 150 ? e.slice(0, 150) + '…' : e}</div>
      <div class="t3">👤 Responsável: <b>${resp(r)}</b></div></div>`;
  }).join('') + (vis.length > 5 ? `<div class="toast mais" style="--c:var(--slate-600)">+ ${vis.length - 5} alerta(s) — veja em Detalhamento dos Processos</div>` : '');
  $$('#toasts .x').forEach(b => b.onclick = () => { const k = decodeURIComponent(b.parentElement.dataset.k); dis[k] = Date.now(); localStorage.setItem('optimusDismiss', JSON.stringify(dis)); render(); });
  let novo = false; vis.forEach(r => { if (!seen[key(r)]) { seen[key(r)] = 1; novo = true; } });
  if (novo) { sessionStorage.setItem('optimusSeen', JSON.stringify(seen)); beep(); }
}
function beep() {
  try { const a = new (window.AudioContext || window.webkitAudioContext)(), o = a.createOscillator(), g = a.createGain(); o.connect(g); g.connect(a.destination); o.type = 'square'; o.frequency.value = 880; g.gain.value = 0.05; o.start(); setTimeout(() => { o.stop(); a.close(); }, 350); } catch (e) {}
}
$('#bBell').onclick = () => { localStorage.removeItem('optimusDismiss'); render(); };

$('#last').textContent = 'Último monitoramento registrado: ' + D.gerado;
setInterval(() => { $('#liveClock').textContent = fmtHMS(new Date()); secLeft = Math.max(0, secLeft - 1); $('#cd').textContent = secLeft + 's'; if (secLeft === 0 && !$('#dWa').open) location.reload(); }, 1000);
$('#bWa').onclick = () => {
  $('#waTxt').value = D.texto;
  const lk = p => `https://web.whatsapp.com/send?phone=${p.numero}&text=${encodeURIComponent($('#waTxt').value)}`;
  $('#waBtns').innerHTML = D.destinos.map((p, i) => `<button class="g" data-i="${i}">Abrir: ${p.nome}</button>`).join('') + '<button id="bCp">Copiar texto</button>';
  $$('#waBtns [data-i]').forEach(x => x.onclick = () => window.open(lk(D.destinos[x.dataset.i]), 'wa' + x.dataset.i));
  $('#bCp').onclick = () => navigator.clipboard.writeText($('#waTxt').value);
  $('#dWa').showModal();
};
$('#bWc').onclick = () => $('#dWa').close();
render();
</script>
</body>
</html>'''


def atualizar_snapshots_hora(res):
    hoje_str = date.today().isoformat()
    dados_snaps = json.loads(SNAPS.read_text(encoding='utf-8')) if SNAPS.exists() else {}
    if dados_snaps.get('dia') != hoje_str:
        dados_snaps = {'dia': hoje_str, 'horas': {}}

    hora_chave = f"{datetime.now().hour:02d}:00"
    dados_snaps['horas'][hora_chave] = {
        'salvo_as': datetime.now().strftime('%H:%M:%S'),
        'itens': res
    }
    SNAPS.write_text(json.dumps(dados_snaps, ensure_ascii=False), encoding='utf-8')
    return dados_snaps['horas']


def gerar_dashboard(cfg, res):
    hist = json.loads(HIST.read_text(encoding='utf-8')) if HIST.exists() else []
    equipe = cfg.get('equipe') or EQUIPE_PADRAO
    ativos = [x for x in res if x['status'] != 'nd']
    c = lambda s: sum(x['status'] == s for x in ativos)
    agora_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    eq = {}
    for x in ativos:
        for p in [q.strip() for q in re.split(r'\s+e\s+|,', x.get('responsavel') or '') if q.strip()]:
            if p in equipe:
                e = eq.setdefault(p, [0, 0]); e[1] += 1; e[0] += (x['status'] == 'ok')
    novo = dict(dia=date.today().isoformat(), eq=eq, h=datetime.now().strftime('%H:%M:%S'), ok=c('ok'), warn=c('warn'), err=c('err'),
                bi_ok=sum(1 for x in ativos if x.get('categoria') == 'pbix' and x['status'] == 'ok'),
                bi_err=sum(1 for x in ativos if x.get('categoria') == 'pbix' and x['status'] == 'err'))
    # 1 registro por hora (o último da hora): mantém o dia inteiro e vários dias no histórico
    if hist and hist[-1].get('dia') == novo['dia'] and hist[-1].get('h', '')[:2] == novo['h'][:2]:
        hist[-1] = novo
    else:
        hist.append(novo)
    hist = hist[-720:]
    HIST.write_text(json.dumps(hist, ensure_ascii=False), encoding='utf-8')

    snaps_hoje = atualizar_snapshots_hora(res)
    incid = atualizar_incidentes(res)
    bi_ev = atualizar_bi_eventos(res)
    dados = dict(
        hoje=date.today().isoformat(),
        equipe=equipe,
        inc=incid.get('resolvidos', [])[-300:],
        inc_abertos=incid.get('abertos', {}),
        bi_ev=bi_ev,
        gerado=agora_str,
        intervalo=cfg.get('intervalo_seg', 60),
        r=res,
        hist=hist,
        snaps=snaps_hoje,
        texto=resumo(res),
        destinos=cfg['whatsapp']['destinos']
    )
    conteudo_html = HTML.replace('__VERSAO__', VERSAO_PAINEL).replace('__DATA__', json.dumps(dados, ensure_ascii=False))

    destino = AQUI / cfg.get('dashboard_html', 'dashboard.html')
    destino.write_text(conteudo_html, encoding='utf-8')
    for arq_extra in ('monitor_automacoes.html', 'monitor_automacoes 1.html', 'monitor_automacoes 1_2.html'):
        p_extra = AQUI / arq_extra
        if p_extra.exists():
            try:
                p_extra.write_text(conteudo_html, encoding='utf-8')
            except Exception:
                pass

    return destino, snaps_hoje, hist


# ==============================================================================
# 8. WHATSAPP A CADA 6 HORAS (360 MIN) - FECHANDO E REINICIANDO POR CONTATO
# ==============================================================================
ICON = {'ok': '✅', 'err': '❌', 'warn': '⚠️', 'run': '🔄', 'nd': '➖'}


def _cut(s, n):
    s = ' '.join(str(s).split())
    return s if len(s) <= n else s[:n - 3] + '...'


def _motivo_curto(msg):
    m = str(msg)
    m = re.sub(r'(\d\d)/(\d\d)/\d{4} (\d\d:\d\d):\d\d', r'\1/\2 \3', m)
    m = re.sub(r'(\d\d/\d\d) (\d\d:\d\d):\d\d', r'\1 \2', m)
    m = re.sub(r'ATRASADO: esperada atualização do horário das (\d\d:\d\d).*?Última: (.+)$', r'Esperado \1; ultima \2', m)
    m = re.sub(r'ATRASADO: sem atualização hoje \(prazo (\S+)\)\. Última: (.+)$', r'Sem atualizar hoje (prazo \1); ultima \2', m)
    m = re.sub(r'NÃO BAIXOU: (.+?) tem (\d+) de (\d+) arquivos hoje', r'Nao baixou \1: \2/\3', m)
    m = re.sub(r'Log do mês não encontrado: .*', 'Log do mes nao encontrado', m)
    m = re.sub(r'Pasta não encontrada: .*', 'Pasta do dia nao encontrada', m)
    return m


def resumo(res, titulo='Monitoramento OptimusBI', max_chars=950):
    """Resumo em formato de planilha (texto monoespaçado, sem emojis), cabendo na legenda de uma imagem."""
    ativos = [x for x in res if x['status'] != 'nd']
    c = lambda s: sum(x['status'] == s for x in ativos)
    bis = [x for x in ativos if x.get('categoria') == 'pbix']
    saude = round(100 * c('ok') / len(ativos)) if ativos else 0
    eq = {}
    for x in ativos:
        for p in [q.strip() for q in re.split(r'\s+e\s+|,', x.get('responsavel') or '') if q.strip()]:
            e = eq.setdefault(p, [0, 0])
            e[1] += 1
            e[0] += (x['status'] == 'ok')
    falhas = sorted([x for x in ativos if x['status'] == 'err'], key=lambda x: (x['prio'], x['nome']))
    alertas = sorted([x for x in ativos if x['status'] == 'warn'], key=lambda x: (x['prio'], x['nome']))

    def montar(n_falhas, n_alertas):
        L = [f"*{titulo}* {datetime.now():%d/%m/%Y %H:%M}", '```',
             f"Saude geral : {saude}%",
             f"OK/Alerta/Falha : {c('ok')}/{c('warn')}/{c('err')}",
             f"BIs em dia  : {sum(x['status'] == 'ok' for x in bis)}/{len(bis)}", '',
             'RESP.      OK/TOT   %']
        for p, (o, t) in sorted(eq.items()):
            L.append(f"{p[:10]:<10} {o:>2}/{t:<4} {round(100 * o / t):>3}%")
        if falhas:
            L += ['', f"FALHAS ({len(falhas)})"]
            for x in falhas[:n_falhas]:
                nome = x['nome'].replace('PBIX · ', '') + (' (pbix)' if x.get('categoria') == 'pbix' else '')
                L.append(f"{x['prio']} {_cut(nome, 30)} | {_cut(x.get('responsavel') or '-', 10)}")
                L.append(f"   {_cut(_motivo_curto(x['erros'][0]) if x['erros'] else '-', 54)}")
            if len(falhas) > n_falhas:
                L.append(f"   (+{len(falhas) - n_falhas} no painel)")
        if alertas and n_alertas:
            L += ['', f"ALERTAS ({len(alertas)}): " + _cut(', '.join(x['nome'].replace('PBIX · ', '') for x in alertas[:n_alertas]), 150)]
        L.append('```')
        return '\n'.join(L)

    for nf, na in [(12, 6), (8, 4), (6, 3), (4, 2), (3, 0), (2, 0), (1, 0), (0, 0)]:
        t = montar(nf, na)
        if len(t) <= max_chars:
            return t
    return t


def _achar_edge(cfg):
    cands = [cfg['whatsapp'].get('edge_exe', ''), r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
             r'C:\Program Files\Microsoft\Edge\Application\msedge.exe', os.path.expandvars(r'%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe')]
    return next((c for c in cands if c and os.path.exists(c)), None)


def capturar_visao_geral(cfg):
    """Tira um print da aba Visão Geral do dashboard (Edge headless, perfil temporário)."""
    try:
        from selenium import webdriver
        from selenium.webdriver.edge.options import Options
        html = AQUI / cfg.get('dashboard_html', 'dashboard.html')
        if not html.exists():
            return None
        op = Options()
        for a in ('--headless=new', '--hide-scrollbars', '--force-device-scale-factor=1', '--window-size=1500,1800', '--disable-gpu',
                  f"--user-data-dir={AQUI / 'edge_print_tmp'}"):
            op.add_argument(a)
        exe = _achar_edge(cfg)
        if exe:
            op.binary_location = exe
        drv = webdriver.Edge(options=op)
        try:
            drv.get(html.as_uri())
            time.sleep(6)  # Chart.js + barras de evolução
            drv.execute_script("var t=document.getElementById('toasts'); if(t) t.style.display='none';")
            altura = min(3400, int(drv.execute_script('return document.documentElement.scrollHeight')) + 20)
            drv.set_window_size(1500, altura)
            time.sleep(2)
            img = AQUI / 'visao_geral.png'
            drv.save_screenshot(str(img))
            log.info('Print da Visão Geral salvo em %s', img)
            return img
        finally:
            drv.quit()
    except Exception as e:
        log.error('Falha ao capturar print da Visão Geral: %s', str(e)[:200])
        return None


def porta_aberta(p):
    import socket
    with socket.socket() as s:
        s.settimeout(1)
        return s.connect_ex(('127.0.0.1', p)) == 0


def encerrar_edge_porta(porta):
    if os.name == 'nt':
        cmd = f'powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object {{ $_.Name -eq \'msedge.exe\' -and $_.CommandLine -like \'*{porta}*\' }} | Invoke-CimMethod -MethodName Terminate" >nul 2>&1'
        subprocess.run(cmd, shell=True)
    time.sleep(2)


def abrir_edge_limpo(cfg):
    w = cfg['whatsapp']
    p = w['porta_edge']
    cands = [
        w.get('edge_exe', ''),
        r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
        r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
        os.path.expandvars(r'%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe')
    ]
    exe = next((c for c in cands if c and os.path.exists(c)), None)
    if not exe:
        log.error('msedge.exe não encontrado.')
        return None

    perfil = w.get('perfil_edge') or r'C:\edge_whatsapp_monitor'
    proc = subprocess.Popen([
        exe,
        f'--remote-debugging-port={p}',
        f'--user-data-dir={perfil}',
        '--no-first-run',
        'https://web.whatsapp.com'
    ])
    for _ in range(40):
        if porta_aberta(p):
            time.sleep(4)
            return proc
        time.sleep(1)
    return None


def _enviar_imagem(drv, caminho_img, legenda):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    icones = "@data-icon='plus' or @data-icon='attach-menu-plus' or @data-icon='clip'"
    WebDriverWait(drv, 30).until(EC.element_to_be_clickable((By.XPATH, f"//footer//span[{icones}]"))).click()
    time.sleep(1)
    inp = WebDriverWait(drv, 15).until(EC.presence_of_element_located((By.XPATH, "//input[@type='file' and contains(@accept,'image')]")))
    inp.send_keys(str(caminho_img))
    time.sleep(4)
    caixa = drv.find_elements(By.XPATH, "//div[@contenteditable='true']")[-1]  # campo de legenda da pré-visualização
    caixa.click()
    for i, linha in enumerate(legenda.split('\n')):
        if i:
            caixa.send_keys(Keys.SHIFT, Keys.ENTER)
        caixa.send_keys(linha)
    time.sleep(1)
    WebDriverWait(drv, 15).until(EC.element_to_be_clickable(
        (By.XPATH, "//span[@data-icon='send' or @data-icon='wds-ic-send-filled']/ancestor::*[@role='button' or self::button][1]"))).click()
    time.sleep(8)


def enviar_whatsapp(cfg, texto, imagem=None):
    """1 mensagem por contato: print da Visão Geral com o resumo na legenda (ou só texto, se o print falhar)."""
    from selenium import webdriver
    from selenium.webdriver.edge.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    w = cfg['whatsapp']
    porta = w['porta_edge']
    destinos = w.get('destinos', [])
    ok_all = True
    xp_caixa = "//footer//div[@contenteditable='true']"

    for idx, d in enumerate(destinos, start=1):
        proc = None
        try:
            log.info('[%s/%s] Iniciando janela do navegador para %s...', idx, len(destinos), d['nome'])
            encerrar_edge_porta(porta)
            proc = abrir_edge_limpo(cfg)
            if not proc:
                ok_all = False
                continue
            op = Options()
            op.debugger_address = f"127.0.0.1:{porta}"
            drv = webdriver.Edge(options=op)

            enviado = False
            if imagem and Path(imagem).exists():
                try:
                    drv.get(f"https://web.whatsapp.com/send?phone={d['numero']}")
                    WebDriverWait(drv, 90).until(EC.presence_of_element_located((By.XPATH, xp_caixa)))
                    time.sleep(2)
                    _enviar_imagem(drv, imagem, texto)
                    enviado = True
                    log.info('WhatsApp enviado para %s (print + resumo). Fechando janela...', d['nome'])
                except Exception as e:
                    log.error('Falha ao enviar print para %s (%s). Enviando só texto.', d['nome'], str(e)[:160])
            if not enviado:
                drv.get(f"https://web.whatsapp.com/send?phone={d['numero']}&text={quote(texto)}")
                caixa = WebDriverWait(drv, 90).until(EC.presence_of_element_located((By.XPATH, xp_caixa)))
                time.sleep(2)
                caixa.send_keys(Keys.ENTER)
                time.sleep(5)
                log.info('WhatsApp enviado para %s (texto). Fechando janela...', d['nome'])
        except Exception as e:
            ok_all = False
            log.error('Falha WhatsApp %s: %s', d['nome'], str(e)[:200])
        finally:
            if proc:
                try:
                    proc.terminate()
                except Exception:
                    pass
            encerrar_edge_porta(porta)

    return ok_all


def rotina_whatsapp(cfg, res, enviar):
    """Envia a atualização no WhatsApp somente a cada 6 horas (360 minutos)."""
    w = cfg['whatsapp']
    if not (enviar and w.get('ativo')):
        return
    est = json.loads(ESTADO.read_text(encoding='utf-8')) if ESTADO.exists() else {}
    agora = datetime.now()
    freq_min = int(w.get('resumo_a_cada_min', 360))  # 360 min = 6 horas
    ultimo = datetime.fromisoformat(est['ultimo_resumo']) if est.get('ultimo_resumo') else None
    ativos = [x for x in res if x['status'] != 'nd']
    erros = sorted(f"{x['nome']}|{e}" for x in ativos for e in x['erros'])
    sig = '\n'.join(erros)

    if ultimo is None or (agora - ultimo) >= timedelta(minutes=freq_min):
        titulo = 'ALERTA - Monitoramento OptimusBI' if erros else 'Monitoramento OptimusBI'
        est['ultimo_resumo'] = agora.isoformat()
        ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')
        imagem = capturar_visao_geral(cfg)
        if enviar_whatsapp(cfg, resumo(res, titulo), imagem):
            est['ultima_assinatura'] = sig
            ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')


# ==============================================================================
# 9. SERVIDOR WEB EMBUTIDO E CICLO PRINCIPAL
# ==============================================================================
ULTIMO_ESTADO_GLOBAL = {'cfg': None, 'res': [], 'snaps': {}, 'hist': []}


# ==============================================================================
# 8B. INCIDENTES (TEMPO DE RESOLUÇÃO) E ATUALIZAÇÕES DE BI POR HORA
# ==============================================================================
INCID = AQUI / 'incidentes.json'
BIEV = AQUI / 'bi_eventos.json'
EQUIPE_PADRAO = ['Karine', 'Agenor', 'Miller']


def _ler_json(p, padrao):
    try:
        return json.loads(Path(p).read_text(encoding='utf-8')) if Path(p).exists() else padrao
    except Exception:
        return padrao


def atualizar_incidentes(res):
    """Abre um incidente quando o item entra em ERRO e fecha (com o tempo de resolução em minutos) quando sai do erro."""
    agora = datetime.now()
    d = _ler_json(INCID, {})
    abertos = d.get('abertos') or {}
    resolvidos = d.get('resolvidos') or []
    for x in res:
        nome, st = x['nome'], x.get('status')
        if st == 'err':
            abertos.setdefault(nome, dict(
                inicio=agora.isoformat(timespec='seconds'), prio=x.get('prio', ''), resp=x.get('responsavel', ''),
                erro=' '.join(str((x.get('erros') or [''])[0]).split())[:220]))
        elif st == 'nd':
            abertos.pop(nome, None)
        elif nome in abertos:
            a = abertos.pop(nome)
            ini = datetime.fromisoformat(a['inicio'])
            resolvidos.append(dict(
                nome=nome, prio=a.get('prio', ''), resp=a.get('resp', ''), inicio=a['inicio'],
                fim=agora.isoformat(timespec='seconds'), min=round((agora - ini).total_seconds() / 60.0, 1),
                dia=agora.date().isoformat(), h=f'{agora.hour:02d}'))
    d = dict(abertos=abertos, resolvidos=resolvidos[-1000:])
    try:
        INCID.write_text(json.dumps(d, ensure_ascii=False), encoding='utf-8')
    except Exception:
        log.error('Falha ao gravar incidentes\n%s', traceback.format_exc())
    return d


def atualizar_bi_eventos(res):
    """Registra cada atualização de BI do dia (detectada pela mudança da data de modificação do .pbix)."""
    hoje = date.today().isoformat()
    d = _ler_json(BIEV, {})
    if d.get('dia') != hoje:
        d = dict(dia=hoje, ult=d.get('ult', {}), eventos=[])
    ult, ev = d['ult'], d['eventos']
    for x in res:
        p = x.get('pbix')
        if x.get('categoria') != 'pbix' or not p or not p.get('ultima_iso'):
            continue
        nome, iso = x['nome'].replace('PBIX · ', ''), p['ultima_iso']
        if ult.get(nome) != iso:
            if iso[:10] == hoje and not any(e['bi'] == nome and e['t'] == iso for e in ev):
                ev.append(dict(bi=nome, t=iso, resp=x.get('responsavel', '')))
            ult[nome] = iso
    try:
        BIEV.write_text(json.dumps(d, ensure_ascii=False), encoding='utf-8')
    except Exception:
        log.error('Falha ao gravar eventos de BI\n%s', traceback.format_exc())
    return d['eventos']


# ==============================================================================
# 8C. PDF DE SAÚDE DA EQUIPE (GERADO A CADA 2 HORAS)
# ==============================================================================
def _donos(x):
    return [q.strip() for q in re.split(r'\s+e\s+|,', x.get('responsavel') or '') if q.strip()]


def _t(s):
    """Texto seguro para o PDF (fontes padrão = Latin-1/cp1252)."""
    return str(s if s is not None else '').encode('cp1252', 'replace').decode('cp1252')


def _grafico_pdf(tipo, caminho, hist, inc, bi_ev, equipe, hoje, agora):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    AZUL, VERM, LAR = '#0339A6', '#F21D2F', '#F24405'
    cores_p = {'Karine': '#F24405', 'Agenor': '#F2B705', 'Miller': '#122A44'}
    fig, ax = plt.subplots(figsize=(11.2, 5.3 if tipo == 'bi' else 2.8), dpi=150)
    por_hora = {}
    for h in hist:
        if h.get('dia') == hoje:
            por_hora[int(h['h'][:2])] = h
    pct = lambda x: round(100 * x['ok'] / max(1, x['ok'] + x['warn'] + x['err']))

    def vazio(msg):
        ax.text(.5, .5, msg, ha='center', va='center', transform=ax.transAxes, color='#64748B')
        ax.set_xticks([]); ax.set_yticks([])

    if tipo == 'equipe':
        hs = sorted(por_hora)
        if not hs:
            vazio('Ainda sem leituras registradas hoje.')
        else:
            ax.plot(hs, [pct(por_hora[h]) for h in hs], color=AZUL, lw=3.2, marker='o', label='Equipe (geral)')
            for n in equipe:
                ys = []
                for h in hs:
                    e = (por_hora[h].get('eq') or {}).get(n)
                    ys.append(round(100 * e[0] / e[1]) if e and e[1] else None)
                pts = [(h, y) for h, y in zip(hs, ys) if y is not None]
                if pts:
                    ax.plot([p[0] for p in pts], [p[1] for p in pts], color=cores_p.get(n, '#94A3B8'), lw=1.8, marker='o', ms=3, label=n)
            ax.set_ylim(0, 105); ax.set_ylabel('% feito (itens em dia)'); ax.set_xlabel('hora do dia')
            ax.set_xticks(hs); ax.set_xticklabels([f'{h:02d}h' for h in hs])
            ax.legend(loc='lower left', ncol=4, fontsize=8, frameon=False)
    elif tipo == 'bi':
        ev = [e for e in bi_ev if e['t'][:10] == hoje]
        if not ev:
            vazio('Nenhuma atualização de BI registrada hoje até o momento.')
        else:
            maxh = max([1, agora.hour] + [int(e['t'][11:13]) for e in ev])
            horas = list(range(maxh + 1))
            nomes = sorted({e['bi'] for e in ev})
            series = [[sum(1 for e in ev if e['bi'] == n and int(e['t'][11:13]) == h) for h in horas] for n in nomes]
            cmap = plt.get_cmap('hsv')
            cores = [cmap(i / max(1, len(nomes))) for i in range(len(nomes))]
            ax.stackplot(horas, *series, labels=nomes, colors=cores, alpha=.8, edgecolor='white', linewidth=.4)
            ax.set_xlim(0, maxh); ax.set_xticks(horas); ax.set_xticklabels([f'{h:02d}h' for h in horas], fontsize=7)
            ax.set_ylabel('atualizações de BI na hora'); ax.set_xlabel('hora do dia')
            ax.yaxis.get_major_locator().set_params(integer=True)
            ax.legend(loc='upper center', bbox_to_anchor=(.5, -.16), ncol=3, fontsize=6.8, frameon=False)
    else:  # falhas x tempo de resolução (duas linhas)
        res_h = {}
        for i in inc.get('resolvidos', []):
            if i.get('dia') == hoje:
                res_h.setdefault(int(i['h']), []).append(i['min'])
        hs = sorted(set(por_hora) | set(res_h))
        if not hs:
            vazio('Ainda sem dados de falhas hoje.')
        else:
            pf = [(h, por_hora[h]['err']) for h in hs if h in por_hora]
            ax.plot([p[0] for p in pf], [p[1] for p in pf], color=VERM, lw=2.4, marker='o', label='Falhas (itens em erro)')
            ax.set_ylabel('falhas', color=VERM); ax.set_ylim(bottom=0)
            ax.yaxis.get_major_locator().set_params(integer=True)
            ax2 = ax.twinx()
            pr = [(h, sum(res_h[h]) / len(res_h[h])) for h in hs if h in res_h]
            if pr:
                ax2.plot([p[0] for p in pr], [p[1] for p in pr], color=AZUL, lw=2.4, marker='s', label='Tempo de resolução (min, média)')
            ax2.set_ylabel('minutos', color=AZUL); ax2.set_ylim(bottom=0)
            ax.set_xticks(hs); ax.set_xticklabels([f'{h:02d}h' for h in hs]); ax.set_xlabel('hora do dia')
            l1, n1 = ax.get_legend_handles_labels(); l2, n2 = ax2.get_legend_handles_labels()
            ax.legend(l1 + l2, n1 + n2, loc='upper left', fontsize=8, frameon=False)
    for s in ('top', 'right'):
        ax.spines[s].set_visible(False)
    ax.grid(axis='y', alpha=.25)
    fig.tight_layout()
    fig.savefig(caminho, bbox_inches='tight')
    plt.close(fig)


def gerar_pdf_saude(cfg, res):
    """Gera o PDF 'Saúde da Equipe OptimusBI' (rotinas executadas + avanço do que é monitorado) e devolve o caminho."""
    import tempfile
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak,
                                    KeepTogether, Flowable, CondPageBreak)

    agora = datetime.now()
    hoje = agora.date().isoformat()
    equipe = cfg.get('equipe') or EQUIPE_PADRAO
    hist = _ler_json(HIST, [])
    inc = _ler_json(INCID, {'abertos': {}, 'resolvidos': []})
    bi_ev = (_ler_json(BIEV, {}).get('eventos')) or []
    janela_h = int(cfg.get('intervalo_pdf_horas', 2))
    desde = agora - timedelta(hours=janela_h)

    AZUL, AZ_ESC, CINZA = colors.HexColor('#0339A6'), colors.HexColor('#122A44'), colors.HexColor('#64748B')
    COR = {'ok': colors.HexColor('#0339A6'), 'warn': colors.HexColor('#B8860B'), 'err': colors.HexColor('#F21D2F'), 'nd': CINZA}
    TXT = {'ok': 'OK', 'warn': 'ALERTA', 'err': 'ERRO', 'nd': 'EM PAUSA'}

    st_n = ParagraphStyle('n', fontName='Helvetica', fontSize=7.4, leading=9)
    st_b = ParagraphStyle('b', parent=st_n, fontName='Helvetica-Bold')
    st_h = ParagraphStyle('h', parent=st_n, fontName='Helvetica-Bold', textColor=colors.white)
    st_t = ParagraphStyle('t', fontName='Helvetica-Bold', fontSize=17, leading=20, textColor=AZ_ESC)
    st_s = ParagraphStyle('s', fontName='Helvetica', fontSize=8.5, leading=11, textColor=CINZA)
    st_2 = ParagraphStyle('2', fontName='Helvetica-Bold', fontSize=11.5, leading=14, textColor=AZUL, spaceBefore=8, spaceAfter=4)
    P = lambda s, e=st_n: Paragraph(escape(_t(s)), e)

    class Barra(Flowable):
        def __init__(self, pct, w=3.6 * cm, h=.32 * cm):
            super().__init__(); self.pct, self.w, self.h = pct, w, h
            self.width, self.height = w, h

        def draw(self):
            cor = COR['ok'] if self.pct >= 90 else colors.HexColor('#F2B705') if self.pct >= 70 else COR['err']
            self.canv.setFillColor(colors.HexColor('#E2E8F0')); self.canv.roundRect(0, 0, self.w, self.h, 2, stroke=0, fill=1)
            self.canv.setFillColor(cor); self.canv.roundRect(0, 0, max(2, self.w * self.pct / 100.0), self.h, 2, stroke=0, fill=1)

    def tabela(dados, larg, extra=None):
        t = Table(dados, colWidths=larg, repeatRows=1, hAlign='LEFT')
        est = [('BACKGROUND', (0, 0), (-1, 0), AZUL), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
               ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F1F5F9')]),
               ('GRID', (0, 0), (-1, -1), .25, colors.HexColor('#CBD5E1')),
               ('TOPPADDING', (0, 0), (-1, -1), 2.2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2.2)]
        t.setStyle(TableStyle(est + (extra or [])))
        return t

    def stat(s, extra=''):
        return Paragraph(f'<font color="{COR.get(s, CINZA).hexval().replace("0x", "#")}"><b>{TXT.get(s, s)}</b></font>' + escape(_t(extra)), st_n)

    ativos = [x for x in res if x.get('status') != 'nd']
    n = lambda lst, s: sum(1 for x in lst if x['status'] == s)
    bis = [x for x in res if x.get('categoria') == 'pbix' and not x.get('pausa')]
    logs_it = next((x for x in res if x.get('categoria') == 'logs'), None)
    saude = round(100 * n(ativos, 'ok') / len(ativos)) if ativos else 0

    # ---- gráficos (PNG temporários) ----
    tmp = Path(tempfile.mkdtemp(prefix='optimus_pdf_'))
    imgs = {}
    for k in ('equipe', 'bi', 'falhas'):
        try:
            _grafico_pdf(k, str(tmp / f'{k}.png'), hist, inc, bi_ev, equipe, hoje, agora)
            imgs[k] = tmp / f'{k}.png'
        except Exception:
            log.error('Falha ao desenhar gráfico %s do PDF\n%s', k, traceback.format_exc())

    def img(k, w=25.2 * cm):
        if k not in imgs:
            return P('(gráfico indisponível)')
        from reportlab.lib.utils import ImageReader
        iw, ih = ImageReader(str(imgs[k])).getSize()
        return Image(str(imgs[k]), width=w, height=w * ih / iw)

    logo = None
    lp = AQUI / 'logo.png'
    if lp.exists():
        try:
            from PIL import Image as PI
            im = PI.open(lp).convert('RGB'); im.thumbnail((360, 360)); logo = tmp / 'logo.jpg'; im.save(logo, quality=88)
        except Exception:
            logo = None

    story = []
    cab = [[Image(str(logo), width=1.9 * cm, height=1.9 * cm) if logo else '',
            [Paragraph('Saúde da Equipe OptimusBI', st_t),
             Paragraph(escape(_t(f'Rotinas executadas e avanço do que é monitorado · gerado em {agora:%d/%m/%Y às %H:%M} · '
                                  f'ciclo de {janela_h} em {janela_h} horas · janela de referência: desde {desde:%H:%M}')), st_s)]]]
    tc = Table(cab, colWidths=[2.4 * cm, 23 * cm]); tc.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('LEFTPADDING', (0, 0), (-1, -1), 0)]))
    story += [tc, Spacer(1, 6)]

    # ---- KPIs ----
    bis_ok = sum(1 for x in bis if x['status'] == 'ok')
    kp = [('Itens monitorados', len(ativos), AZUL), ('Em dia (OK)', n(ativos, 'ok'), AZUL), ('Alertas', n(ativos, 'warn'), colors.HexColor('#B8860B')),
          ('Crítico / Erro', n(ativos, 'err'), COR['err']), ('BIs em dia', f'{bis_ok}/{len(bis)}', AZ_ESC),
          ('Saúde geral', f'{saude}%', COR['ok'] if saude >= 90 else colors.HexColor('#B8860B') if saude >= 70 else COR['err'])]
    cel = [[Paragraph(f'<font size="17" color="{c.hexval().replace("0x", "#")}"><b>{escape(str(v))}</b></font><br/><font size="7.5" color="#64748B">{a}</font>',
                      ParagraphStyle('k', alignment=1, leading=20)) for a, v, c in kp]]
    tk = Table(cel, colWidths=[4.2 * cm] * 6)
    tk.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), .5, colors.HexColor('#CBD5E1')), ('INNERGRID', (0, 0), (-1, -1), .5, colors.HexColor('#CBD5E1')),
                            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')), ('TOPPADDING', (0, 0), (-1, -1), 6), ('BOTTOMPADDING', (0, 0), (-1, -1), 6)]))
    story.append(tk)

    ab_janela = [a for a in inc.get('abertos', {}).values() if a['inicio'] >= desde.isoformat()]
    res_janela = [i for i in inc.get('resolvidos', []) if i['fim'] >= desde.isoformat()]
    ini_janela = ab_janela + [i for i in inc.get('resolvidos', []) if i['inicio'] >= desde.isoformat()]
    ev_janela = [e for e in bi_ev if e['t'] >= desde.isoformat(timespec='seconds')]
    story.append(Paragraph(escape(_t(
        f'Últimas {janela_h} h: {len(ev_janela)} atualização(ões) de BI · {len(ini_janela)} falha(s) aberta(s) · {len(res_janela)} falha(s) resolvida(s) · '
        f'{len(inc.get("abertos", {}))} falha(s) em aberto agora.')), ParagraphStyle('j', parent=st_s, spaceBefore=5)))

    # ---- Saúde da equipe ----
    story.append(Paragraph('Saúde da equipe — % de itens em dia por responsável', st_2))
    linhas = [('Equipe OptimusBI (geral)', ativos)] + [(p, [x for x in ativos if p in _donos(x)]) for p in equipe]
    d = [[P('Responsável', st_h), P('Itens', st_h), P('Em dia', st_h), P('Alertas', st_h), P('Erros', st_h), P('% feito', st_h), P('Avanço', st_h)]]
    for nome, lst in linhas:
        pc = round(100 * n(lst, 'ok') / len(lst)) if lst else 0
        d.append([P(nome, st_b), P(len(lst)), P(n(lst, 'ok')), P(n(lst, 'warn')), P(n(lst, 'err')), P(f'{pc}%', st_b), Barra(pc)])
    story.append(tabela(d, [6.2 * cm, 2 * cm, 2 * cm, 2 * cm, 2 * cm, 2.2 * cm, 4.4 * cm]))

    # ---- Avanço por categoria ----
    story.append(Paragraph('Avanço do que é monitorado', st_2))
    cats = [('Fluxos de execução (logs das rotinas)', [x for x in ativos if x.get('categoria') == 'bi']),
            ('Download de bases', [x for x in ativos if x.get('categoria') == 'base']),
            ('BIs (.pbix oficiais)', bis)]
    d = [[P('Categoria', st_h), P('Total', st_h), P('OK', st_h), P('Alertas', st_h), P('Erros', st_h), P('% em dia', st_h), P('Avanço', st_h)]]
    for nome, lst in cats:
        pc = round(100 * n(lst, 'ok') / len(lst)) if lst else 0
        d.append([P(nome, st_b), P(len(lst)), P(n(lst, 'ok')), P(n(lst, 'warn')), P(n(lst, 'err')), P(f'{pc}%', st_b), Barra(pc)])
    if logs_it:
        ls = logs_it.get('logs_detalhe') or []
        pc = round(100 * sum(1 for l in ls if l['status'] == 'ok') / len(ls)) if ls else 0
        d.append([P('Auditoria de logs (Pasta 52)', st_b), P(len(ls)), P(sum(1 for l in ls if l['status'] == 'ok')),
                  P(sum(1 for l in ls if l['status'] == 'warn')), P(sum(1 for l in ls if l['status'] == 'err')), P(f'{pc}%', st_b), Barra(pc)])
    story.append(tabela(d, [6.2 * cm, 2 * cm, 2 * cm, 2 * cm, 2 * cm, 2.2 * cm, 4.4 * cm]))
    story.append(KeepTogether([Paragraph('Evolução da equipe por hora (hoje) — % de feito', st_2), img('equipe')]))

    # ---- Rotinas executadas ----
    story += [CondPageBreak(7 * cm), Paragraph('Rotinas executadas — fluxos de execução e download de bases', st_2)]
    rot = sorted([x for x in ativos if x.get('categoria') in ('bi', 'base')], key=lambda x: (x.get('prio', ''), x['nome']))
    d = [[P(h, st_h) for h in ('Rotina', 'Resp.', 'Prio', 'Tipo', 'Status', 'Exec. hoje', 'OK / Erro', 'Último fim', 'Tempo médio', 'Pastas (arquivos hoje / meta)')]]
    extra = []
    for i, x in enumerate(rot, 1):
        pastas = '; '.join(f"{f['rotulo']} {f['n']}/{f['esperado']}" if f.get('n') is not None else f"{f['rotulo']} (erro)" for f in x.get('pastas', []))
        d.append([P(x['nome'], st_b), P(x.get('responsavel') or '-'), P(x.get('prio', '')), P('Download' if x.get('categoria') == 'base' else 'Fluxo'), stat(x['status']),
                  P('-' if x.get('linhas_hoje') is None else x['linhas_hoje']),
                  P('-' if x.get('ok_hoje') is None else f"{x['ok_hoje']} / {x['err_hoje']}"), P(x.get('ultimo_fim') or '-'),
                  P('-' if x.get('media_min') is None else f"{x['media_min']:.1f} min"), P(pastas or '-')])
    story.append(tabela(d, [5.3 * cm, 1.9 * cm, 1.1 * cm, 1.8 * cm, 1.6 * cm, 1.6 * cm, 1.9 * cm, 1.9 * cm, 2 * cm, 6.1 * cm], extra))
    problemas = [x for x in rot if x['status'] in ('err', 'warn')]
    if problemas:
        story.append(CondPageBreak(4 * cm))
        story.append(Paragraph('Pontos de atenção nas rotinas', st_2))
        d = [[P('Rotina', st_h), P('Resp.', st_h), P('Situação', st_h)]]
        for x in problemas:
            d.append([P(x['nome'], st_b), P(x.get('responsavel') or '-'), P(' | '.join((x.get('erros') or []) + (x.get('avisos') or []))[:420])])
        story.append(tabela(d, [6 * cm, 2.4 * cm, 16.8 * cm]))

    # ---- BIs ----
    story += [CondPageBreak(7 * cm), Paragraph('Monitoramento dos BIs — atualização dos .pbix oficiais', st_2)]
    d = [[P(h, st_h) for h in ('BI', 'Resp.', 'Prio', 'Última atualização', 'Regra esperada', 'Idade (h)', 'Atraso (min)', 'Status')]]
    for x in sorted([x for x in res if x.get('categoria') == 'pbix'], key=lambda x: (x.get('prio', ''), x['nome'])):
        p = x.get('pbix') or {}
        pausa = x.get('pausa')
        d.append([P(x['nome'].replace('PBIX · ', ''), st_b), P(x.get('responsavel') or '-'), P(x.get('prio', '')), P(p.get('ultima') or '-'), P(p.get('esperado') or '-'),
                  P('-' if p.get('idade_h') is None else p['idade_h']), P(p.get('atraso_min') or 0), stat('nd' if pausa else x['status'])])
    story.append(tabela(d, [6.2 * cm, 2.4 * cm, 1.1 * cm, 3.4 * cm, 6.2 * cm, 1.7 * cm, 2 * cm, 2.2 * cm]))
    story.append(KeepTogether([Paragraph('Atualizações de BI por hora (hoje) — área empilhada, uma cor por BI', st_2), img('bi')]))

    # ---- Falhas x resolução ----
    story.append(KeepTogether([Paragraph('Falhas × tempo de resolução por hora (hoje)', st_2), img('falhas')]))
    story.append(CondPageBreak(5 * cm))
    story.append(Paragraph('Falhas em aberto agora', st_2))
    d = [[P(h, st_h) for h in ('Item', 'Resp.', 'Prio', 'Aberta desde', 'Tempo aberta', 'Motivo')]]
    abertos = sorted(inc.get('abertos', {}).items(), key=lambda kv: kv[1]['inicio'])
    for nome, a in abertos:
        ini = datetime.fromisoformat(a['inicio']); mins = int((agora - ini).total_seconds() // 60)
        d.append([P(nome.replace('PBIX · ', ''), st_b), P(a.get('resp') or '-'), P(a.get('prio', '')), P(f'{ini:%d/%m %H:%M}'),
                  P(f'{mins // 60}h{mins % 60:02d}min'), P(a.get('erro', ''))])
    if len(d) == 1:
        d.append([P('Nenhuma falha em aberto no momento.'), '', '', '', '', ''])
    story.append(tabela(d, [6 * cm, 2.2 * cm, 1.1 * cm, 2.6 * cm, 2.4 * cm, 10.9 * cm]))
    story.append(CondPageBreak(5 * cm))
    story.append(Paragraph('Falhas resolvidas hoje', st_2))
    d = [[P(h, st_h) for h in ('Item', 'Resp.', 'Prio', 'Aberta', 'Resolvida', 'Tempo de resolução')]]
    for i in [i for i in inc.get('resolvidos', []) if i.get('dia') == hoje][-15:][::-1]:
        d.append([P(i['nome'].replace('PBIX · ', ''), st_b), P(i.get('resp') or '-'), P(i.get('prio', '')), P(i['inicio'][11:16]), P(i['fim'][11:16]), P(f"{i['min']:.0f} min")])
    if len(d) == 1:
        d.append([P('Nenhuma falha resolvida hoje até o momento.'), '', '', '', '', ''])
    story.append(tabela(d, [8 * cm, 3 * cm, 1.4 * cm, 3 * cm, 3 * cm, 5.8 * cm]))

    def rodape(c, doc):
        c.saveState(); c.setFont('Helvetica', 7); c.setFillColor(CINZA)
        c.drawString(1.2 * cm, .8 * cm, _t(f'OptimusBI Squad · Monitoramento {VERSAO_PAINEL.split(" ·")[0]} · {agora:%d/%m/%Y %H:%M}'))
        c.drawRightString(landscape(A4)[0] - 1.2 * cm, .8 * cm, f'Página {doc.page}')
        c.restoreState()

    nome_pdf = f'Saude_Equipe_OptimusBI_{agora:%Y-%m-%d_%H%M}.pdf'
    pastas_saida = [AQUI / 'pdf_saude_equipe'] + [Path(p) for p in cfg.get('pastas_pdf_saude', []) if p]
    principal = None
    for i, pasta in enumerate(pastas_saida):
        try:
            pasta.mkdir(parents=True, exist_ok=True)
            destino = pasta / nome_pdf
            if principal is None:
                SimpleDocTemplate(str(destino), pagesize=landscape(A4), leftMargin=1.2 * cm, rightMargin=1.2 * cm, topMargin=1.1 * cm,
                                  bottomMargin=1.4 * cm, title='Saúde da Equipe OptimusBI', author='OptimusBI Squad').build(story, onFirstPage=rodape, onLaterPages=rodape)
                principal = destino
            else:
                import shutil
                shutil.copy2(principal, destino)
        except Exception:
            log.error('Não foi possível salvar o PDF em %s\n%s', pasta, traceback.format_exc())
    try:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
    except Exception:
        pass
    if principal:
        log.info('PDF de saúde da equipe gerado: %s', principal)
    return principal


def rotina_pdf_2h(cfg, res, forcar=False):
    """Gera o PDF a cada 2 horas (00h, 02h, 04h ... 22h) — 1 PDF por faixa, mesmo que o monitor tenha ficado parado."""
    agora = datetime.now()
    passo = max(1, int(cfg.get('intervalo_pdf_horas', 2)))
    faixa = f'{date.today().isoformat()}_{agora.hour // passo * passo:02d}'
    est = _ler_json(ESTADO, {})
    if forcar or est.get('ultimo_pdf_faixa') != faixa:
        arq = gerar_pdf_saude(cfg, res)
        if arq:
            est = _ler_json(ESTADO, {})
            est['ultimo_pdf_faixa'] = faixa
            est['ultimo_pdf_arquivo'] = str(arq)
            ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')
            print(f'[OptimusBI] PDF de saúde da equipe salvo em: {arq}')
        return arq
    return None


class PainelHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        return

    def do_GET(self):
        if self.path.startswith('/api/atualizar'):
            ciclo(enviar=False)
            self.send_response(200)
            self.send_header('Content-type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'ok': True, 'hora': datetime.now().strftime('%H:%M:%S')}).encode('utf-8'))
            return

        if self.path.startswith('/api/pdf'):
            cfg = ULTIMO_ESTADO_GLOBAL['cfg'] or sincronizar_config_automatico()
            if not ULTIMO_ESTADO_GLOBAL['res']:
                ciclo(enviar=False)
            arq = gerar_pdf_saude(cfg, ULTIMO_ESTADO_GLOBAL['res'])
            self.send_response(200)
            self.send_header('Content-type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'ok': bool(arq), 'arquivo': str(arq or '')}).encode('utf-8'))
            return

        if self.path.startswith('/api/relatorio'):
            cfg = ULTIMO_ESTADO_GLOBAL['cfg'] or sincronizar_config_automatico()
            arq = gerar_relatorio_xlsx(
                cfg, ULTIMO_ESTADO_GLOBAL['res'], ULTIMO_ESTADO_GLOBAL['snaps'],
                ULTIMO_ESTADO_GLOBAL['hist'], salvar_em_pasta_especifica=cfg.get('pasta_relatorio_diario')
            )
            self.send_response(200)
            self.send_header('Content-type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'ok': True, 'arquivo': str(arq)}).encode('utf-8'))
            return

        if self.path.startswith('/api/vercel'):
            cfg = ULTIMO_ESTADO_GLOBAL['cfg'] or sincronizar_config_automatico()
            arq = gerar_relatorio_xlsx(cfg, ULTIMO_ESTADO_GLOBAL['res'], ULTIMO_ESTADO_GLOBAL['snaps'], ULTIMO_ESTADO_GLOBAL['hist'])
            enviar_xlsx_para_painel_vercel(cfg, arq)
            self.send_response(200)
            self.send_header('Content-type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'ok': True, 'arquivo': str(arq)}).encode('utf-8'))
            return

        if self.path.startswith('/logo.png') and (AQUI / 'logo.png').exists():
            self.send_response(200)
            self.send_header('Content-type', 'image/png')
            self.end_headers()
            self.wfile.write((AQUI / 'logo.png').read_bytes())
            return

        arq_html = AQUI / (ULTIMO_ESTADO_GLOBAL['cfg'] or DEFAULT).get('dashboard_html', 'dashboard.html')
        if not arq_html.exists():
            ciclo(enviar=False)
        self.send_response(200)
        self.send_header('Content-type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(arq_html.read_bytes())


def iniciar_servidor_web(porta=8080):
    try:
        srv = HTTPServer(('127.0.0.1', porta), PainelHandler)
        t = threading.Thread(target=srv.serve_forever, daemon=True)
        t.start()
        return True
    except Exception:
        return False


def verificar_relatorio_final_do_dia(cfg, res, snaps_hoje, hist, forcar=False):
    agora = datetime.now()
    est = json.loads(ESTADO.read_text(encoding='utf-8')) if ESTADO.exists() else {}
    hoje_iso = date.today().isoformat()
    hora_alvo = int(cfg.get('hora_relatorio_diario', 18))

    ja_gerou_hoje = (est.get('ultimo_relatorio_xlsx') == hoje_iso)
    if forcar or (agora.hour >= hora_alvo and not ja_gerou_hoje):
        arq_salvo = gerar_relatorio_xlsx(
            cfg, res, snaps_hoje, hist,
            salvar_em_pasta_especifica=cfg.get('pasta_relatorio_diario', DEFAULT['pasta_relatorio_diario'])
        )
        est['ultimo_relatorio_xlsx'] = hoje_iso
        ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')
        print(f"[OptimusBI] Relatório diário .xlsx salvo em: {arq_salvo}")


def ciclo(enviar, forcar_relatorio=False, forcar_vercel=False, forcar_pdf=False):
    cfg = sincronizar_config_automatico()
    res = []
    for m in cfg['monitores']:
        try:
            res.append(avaliar(m, cfg))
        except Exception:
            log.error('Erro avaliando %s\n%s', m['nome'], traceback.format_exc())
            res.append(dict(
                nome=m['nome'], categoria=m.get('categoria', 'bi'), prio=m['prio'], status='err',
                erros=['Falha interna na verificação'], avisos=[], infos=[], pastas=[],
                quando=datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                linhas_hoje=None, ok_hoje=None, err_hoje=None, media_min=None, media_arq_min=None, ultimo_fim=None
            ))

    _por_nome = {m['nome']: m for m in cfg['monitores']}
    for x in res:
        mm = _por_nome.get(x['nome'])
        if mm:
            x['responsavel'] = mm.get('responsavel', '')
            x['tipo_proc'] = mm.get('tipo_proc') or ('download' if x.get('categoria') == 'base' else 'fluxo')
    try:
        res.extend(avaliar_pbix(cfg))
        res.append(auditar_logs_52(cfg))
    except Exception:
        log.error('Erro no monitoramento de BIs/logs\n%s', traceback.format_exc())

    dest, snaps_hoje, hist = gerar_dashboard(cfg, res)
    ULTIMO_ESTADO_GLOBAL.update(cfg=cfg, res=res, snaps=snaps_hoje, hist=hist)

    ativos = [x for x in res if x['status'] != 'nd']
    log.info(
        'Ciclo concluído: %s | dashboard: %s',
        {s: sum(x['status'] == s for x in ativos) for s in ('ok', 'warn', 'err', 'run')},
        dest
    )

    try:
        rotina_vercel_30_minutos(cfg, res, snaps_hoje, hist, forcar=forcar_vercel)
    except Exception:
        log.error('Falha na rotina Vercel 30m\n%s', traceback.format_exc())

    try:
        verificar_relatorio_final_do_dia(cfg, res, snaps_hoje, hist, forcar=forcar_relatorio)
    except Exception:
        log.error('Falha ao gerar relatório diário .xlsx\n%s', traceback.format_exc())

    try:
        rotina_pdf_2h(cfg, res, forcar=forcar_pdf)
    except Exception:
        log.error('Falha na rotina do PDF de saúde (2h)\n%s', traceback.format_exc())

    try:
        rotina_whatsapp(cfg, res, enviar)
    except Exception:
        log.error('Falha na rotina WhatsApp\n%s', traceback.format_exc())

    return cfg


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--uma-vez', action='store_true')
    ap.add_argument('--sem-whatsapp', action='store_true')
    ap.add_argument('--testar-whatsapp', action='store_true')
    ap.add_argument('--testar-resumo', action='store_true', help='Gera o resumo real + print da Visão Geral e envia agora')
    ap.add_argument('--so-print', action='store_true', help='Só gera o print visao_geral.png (sem enviar)')
    ap.add_argument('--gerar-relatorio', action='store_true')
    ap.add_argument('--gerar-pdf', action='store_true', help='Gera agora o PDF de saúde da equipe')
    ap.add_argument('--enviar-vercel', action='store_true', help='Força geração do .xlsx e upload imediato na Vercel')
    a = ap.parse_args()

    if a.testar_whatsapp:
        sucesso = enviar_whatsapp(sincronizar_config_automatico(), 'Teste Monitoramento OptimusBI ✅ (1 mensagem por contato)')
        print('Enviado!' if sucesso else 'Falhou - veja monitor.log')
        sys.exit()

    if a.testar_resumo or a.so_print:
        cfg_t = ciclo(enviar=False)
        img = capturar_visao_geral(cfg_t)
        print('Print:', img)
        if a.testar_resumo:
            print('Enviado!' if enviar_whatsapp(cfg_t, resumo(ULTIMO_ESTADO_GLOBAL['res']), img) else 'Falhou - veja monitor.log')
        sys.exit()

    cfg_ini = sincronizar_config_automatico()
    if not a.uma_vez and not a.gerar_relatorio and not a.enviar_vercel and not a.gerar_pdf:
        porta_web = int(cfg_ini.get('porta_servidor_web', 8080))
        if not iniciar_servidor_web(porta_web):
            msg = (f'\n*** ATENÇÃO: a porta {porta_web} já está em uso. Existe OUTRO monitor rodando (provavelmente a versão antiga), '
                   f'que continua servindo e regravando o dashboard antigo. Feche a outra janela do monitor (ou finalize o python.exe '
                   f'no Gerenciador de Tarefas) e abra este de novo. ***\n')
            print(msg); log.error(msg.strip())
            sys.exit(3)
        log.info('Monitoramento OptimusBI %s iniciado (PID %s)', VERSAO_PAINEL, os.getpid())
        if cfg_ini.get('abrir_navegador_ao_iniciar', True):
            webbrowser.open(f'http://localhost:{porta_web}')

    while True:
        ini = time.time()
        try:
            cfg = ciclo(not a.sem_whatsapp, forcar_relatorio=a.gerar_relatorio, forcar_vercel=a.enviar_vercel, forcar_pdf=a.gerar_pdf)
        except Exception:
            log.error('Erro no ciclo\n%s', traceback.format_exc())
            cfg = DEFAULT
        if a.uma_vez or a.gerar_relatorio or a.enviar_vercel or a.gerar_pdf:
            break
        tempo_espera = max(10, int(cfg.get('intervalo_seg', 60)) - (time.time() - ini))
        time.sleep(tempo_espera)