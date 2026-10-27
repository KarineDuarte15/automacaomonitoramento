#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitoramento OptimusBI v6.0 - Completo:
- Limpa processos antigos na porta 8080 e desativa cache do navegador (evita espelhar tela antiga)
- Verifica pastas, BIs (.pbix) e logs a cada 60s
- Alimenta o Teams Desktop no grupo 'TIME BI OPTIMUS'
- Gera PDF executivo da Visão Geral (dashboard + explicação completa)
- Dispara e-mail pelo Outlook no navegador às 07:00 e às 17:00 com o PDF anexado e explicação no corpo
- Ao final do dia (17:00+): entra no Microsoft Planner no navegador e dá check nas atividades
- A cada 30 minutos: gera .xlsx e sobe no painel Vercel
- A cada 6 horas: envia resumo no WhatsApp
"""
import argparse
import base64
import calendar
import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
import traceback
import unicodedata
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
JANELAS_TELE_MIN = [h * 60 - 30 for h in HORARIOS_APRR]


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
# 2. CONFIGURAÇÃO PADRÃO (INCLUINDO TEAMS, OUTLOOK 07H/17H, PLANNER, VERCEL E WA)
# ==============================================================================
DEFAULT = {
    'versao_config': 6,
    'raiz': r'G:\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES',
    'raiz_rede_fallback': r'\\10.1.17.4\Usuarios\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES',
    'pasta_logs': r'1 - RELATORIOS\52 - LOGS PROJETOS',
    'pasta_pbix': r'3 - ATUALIZAÇÃO PBIX\PBIX OFICIAL',
    'hora_limite_manha': '11:00',
    'bi_checar_fim_de_semana': False,
    'logs_ignorar': [],
    'pasta_relatorio_diario': r'G:\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES\7 - BASES DE RELATORIOS\RELATORIO PAINEL MONITORAMENTO DIARIO',
    'pasta_dashboard_acompanhamento': r'C:\Users\agenor.sales\Desktop\dashboard acompanhamento',
    'url_painel_vercel': 'https://acompanhameto-fluxo-diario-de-taref.vercel.app/',
    'intervalo_vercel_min': 30,
    'hora_relatorio_diario': 18,
    'intervalo_seg': 60,
    'porta_servidor_web': 8080,
    'abrir_navegador_ao_iniciar': True,
    'dashboard_html': 'dashboard.html',
    'teams': {
        'ativo': True,
        'grupo': 'TIME BI OPTIMUS',
        'intervalo_min': 120,
        'horarios_fixos': ['07:00', '12:00', '17:00']
    },
    'email_outlook': {
        'ativo': True,
        'horarios': ['07:00', '17:00'],
        'porta_chrome': 9226,
        'perfil_chrome': r'C:\chrome_outlook_planner',
        'destinatarios': [
            'erika.karine@hapvida.com.br'
        ],
        'assunto_prefixo': '[OptimusBI] Visão Geral do Monitoramento'
    },
    'planner': {
        'ativo': True,
        'horario_check': '17:05',
        'url_planner': 'https://planner.cloud.microsoft/',
        'porta_chrome': 9226,
        'perfil_chrome': r'C:\chrome_outlook_planner',
        'palavras_chave_tarefas': []
    },
    'whatsapp': {
        'ativo': True,
        'porta_chrome': 9222,
        'resumo_a_cada_min': 360,
        'horarios_resumo': ['00:00', '06:00', '12:00', '18:00', '23:59'],
        'alerta_imediato_erro': False,
        'perfil_chrome': r'C:\chrome_whatsapp_monitor',
        'chrome_exe': '',
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
                F('T0012 DISPONIBILIDADE', B7 + r'T0012 - ULTIMA BASE DISPONIBILIDADE MENSAL\{YYYY}', 14, '', True, '', conf_mensal=True),
                F('T0012 MARCADA', B7 + r'T0012 - ULTIMA BASE MARCADA MENSAL\{YYYY}', 14, '', True, '', conf_mensal=True)
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
                F('T9016', B7 + 'T9016 - BASE MENSAL' + MM, 5, '', True, '', conf_mensal=False,
                  dias_mes=[1, 5, 10], prefixos=['T9016_', 'NDISPRJ', 'NDIMINAS', 'CLINIPAM', 'CCG'])
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


def _mon(nome):
    return next(x for x in DEFAULT['monitores'] if x['nome'] == nome)


def _bi(nome, busca, prio, regra='diaria', **kw):
    return dict(nome=nome, busca=busca, prio=prio, regra=regra, **kw)


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
    _bi('Acompanhamento de Pagamentos Prestadores', ['pagamentos', 'prestadores'], 'P1', 'ultima', resp='Agenor'),
    _bi('Escala Especial', ['escala', 'especial'], 'P2', resp='Agenor'),
    _bi('Relatório Atraso dos Médicos', ['atraso', 'medicos'], 'P2', resp='Karine'),
    _bi('Relatório de Consultas em Transferência', ['transferencia'], 'P2', resp='Karine'),
    _bi('Relatório de Cancelamento v2', ['cancelamento'], 'P2', resp='Karine'),
    _bi('Relatório de Acompanhamento de Mensageria', ['mensageria'], 'P2', resp='Karine'),
    _bi('Monitoramento de Marcação', ['monitoramento', 'marcacao'], 'P2', 'ate_hora', hora_limite='06:00', tolerancia_min=60, resp='Karine'),
    _bi('Relatório de Cirurgias', ['cirurgias'], 'P2', resp='Karine'),
    _bi('Relatório SAC (mensal)', ['sac'], 'P2', 'mensal', excluir=['nip'], resp='Karine'),
    _bi('Médicos Novos (mensal)', ['medicos', 'novos'], 'P2', 'mensal', resp='Karine'),
    _bi('Relatório de SAC e NIP (desenvolvimento)', ['sac', 'nip'], 'P2', 'horarios', horarios=['07:30', '09:00', '13:00', '22:30'], tolerancia_min=90, resp='Daniel'),
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

for _m in DEFAULT['monitores']:
    _m['responsavel'] = _RESP.get(_m['nome'], '')

DEFAULT['monitores'].append({
    'nome': 'DB T0012 (mensal)', 'categoria': 'base', 'prio': 'P1', 'regra': '', 'log': '', 'min_linhas': 0, 'responsavel': 'Agenor',
    'pastas': [F('DB_T0012 FLUXO PYTHON', B7 + r'DB_T0012\FLUXO PYTHON', 3, '', True, '', conf_mensal=True)]
})
DEFAULT['monitores'].append({
    'nome': 'Indicadores Rede Cred (bases)', 'categoria': 'base', 'prio': 'P1', 'regra': '', 'log': '', 'min_linhas': 0, 'responsavel': 'Karine',
    'pastas': [
        F('CLIENTE_OCULTO HISTORICO', B7 + r'BI_INDICADORES_CRED\CLIENTE_OCULTO\HISTORICO', 1, '', True, '', conf_mensal=True, hora_limite='10:00'),
        F('CLIENTE_OCULTO MENSAL', B7 + r'BI_INDICADORES_CRED\CLIENTE_OCULTO', 1, '', True, '', conf_mensal=True, hora_limite='10:00'),
    ]
})
DEFAULT['monitores'].append({
    'nome': 'Base DGE (produtividade / pagamento)', 'categoria': 'base', 'prio': 'P2', 'regra': '', 'log': '', 'min_linhas': 0, 'responsavel': 'Karine',
    'pastas': [
        F('48 - PRODUTIVIDADE V2 DGE', r'1 - RELATORIOS\48 - RELATORIO DE PRODUTIVIDADE V2\DGE', 1, '', False, '', so_mes=True),
        F('BASES DGE PRODUTIVIDADE', B7 + r'BASES DGE\PRODUTIVIDADE', 1, '', False, '', so_mes=True),
        F('BASES DGE TIPO PAGAMENTO', B7 + r'BASES DGE\TIPO PAGAMENTO', 1, '', False, '', so_mes=True),
    ]
})

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
            if not isinstance(antigo, dict):
                raise ValueError('a raiz da configuração deve ser um objeto JSON')

            def mesclar(base, personalizado):
                for chave, valor in personalizado.items():
                    if isinstance(valor, dict) and isinstance(base.get(chave), dict):
                        mesclar(base[chave], valor)
                    else:
                        base[chave] = valor

            mesclar(cfg_final, antigo)
        except (OSError, json.JSONDecodeError, TypeError, ValueError) as e:
            log.warning('Configuração personalizada inválida; usando padrões: %s', e)
    cfg_final['whatsapp']['resumo_a_cada_min'] = 360
    cfg_final['intervalo_vercel_min'] = 30
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
# 4. REGRAS DE HORÁRIO: APRR E TELE SAÚDE + PBIX + PASTA 52
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
    hoje = agora.date()
    regra = b.get('regra', 'diaria')
    if regra == 'intervalo_horas':
        intervalo_h = float(b.get('intervalo_h', 2))
        idade_h = max(0, (agora - mt).total_seconds() / 3600)
        if idade_h <= intervalo_h:
            return 'ok', f'Atualizado há {idade_h:.1f} hora(s)', f'Atualizar a cada {intervalo_h:g} hora(s)', 0
        atraso = int((idade_h - intervalo_h) * 60)
        return 'err', f'ATRASADO: última atualização há {idade_h:.1f} hora(s) (limite {intervalo_h:g} h)', f'Atualizar a cada {intervalo_h:g} hora(s)', atraso
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
            r['avisos'].append(f"Arquivo .pbix não localizado (busca: {', '.join(b['busca'])}).")
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
# 6. GERAÇÃO DO .XLSX E AUTOMAÇÃO DE UPLOAD NA VERCEL A CADA 30 MINUTOS
# ==============================================================================
def gerar_relatorio_xlsx(cfg, res, snaps_hoje, hist_diario, salvar_em_pasta_especifica=None):
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

    if salvar_em_pasta_especifica:
        try:
            p_rede = Path(salvar_em_pasta_especifica)
            p_rede.mkdir(parents=True, exist_ok=True)
            wb.save(p_rede / nome_arq)
        except Exception as e:
            log.warning('Não foi possível salvar cópia na rede %s: %s', salvar_em_pasta_especifica, e)

    return arq_local_datado


def enviar_xlsx_para_painel_vercel(cfg, caminho_xlsx):
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    url_vercel = cfg.get('url_painel_vercel', DEFAULT['url_painel_vercel'])
    porta_vercel = 9225
    proc = None

    try:
        log.info('Iniciando upload do arquivo %s para o painel Vercel (%s)...', caminho_xlsx.name, url_vercel)
        encerrar_chrome_porta(porta_vercel)
        exe = _achar_chrome(cfg)
        if not exe:
            log.error('chrome.exe não encontrado para subir planilha na Vercel.')
            return False

        perfil_vercel = r'C:\chrome_vercel_monitor'
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
        drv = webdriver.Chrome(options=op)
        drv.get(url_vercel)
        time.sleep(5)

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

        inputs_file = drv.find_elements(By.XPATH, "//input[@type='file']")
        if not inputs_file:
            inp = WebDriverWait(drv, 15).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
            )
        else:
            inp = inputs_file[0]

        drv.execute_script("arguments[0].style.display = 'block'; arguments[0].style.visibility = 'visible';", inp)
        time.sleep(1)
        inp.send_keys(str(caminho_xlsx.resolve()))
        time.sleep(6)

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
        encerrar_chrome_porta(porta_vercel)


def rotina_vercel_30_minutos(cfg, res, snaps_hoje, hist, forcar=False):
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
# 7. NOVAS AUTOMAÇÕES: TEAMS DESKTOP, PDF EXECUTIVO, OUTLOOK WEB E PLANNER
# ==============================================================================
def enviar_teams_desktop(cfg, texto):
    """
    Simula o clique e atalhos na tela do aplicativo Microsoft Teams Desktop,
    entra no grupo 'TIME BI OPTIMUS' e envia o relatório atualizado.
    """
    import pyautogui
    import pyperclip
    from pywinauto import Desktop

    t_cfg = cfg.get('teams', {})
    grupo = t_cfg.get('grupo', 'TIME BI OPTIMUS')

    try:
        log.info('Iniciando envio para o Teams Desktop (Grupo: %s)...', grupo)
        pyautogui.FAILSAFE = False

        janelas = Desktop(backend="uia").windows(title_re=".*Teams.*", visible_only=False)
        if not janelas:
            os.startfile("msteams:")
            time.sleep(6)
            janelas = Desktop(backend="uia").windows(title_re=".*Teams.*", visible_only=False)
            if not janelas:
                log.error('Janela do Teams Desktop não encontrada.')
                return False

        win = janelas[0]
        if win.is_minimized():
            win.restore()
        win.set_focus()
        time.sleep(1.5)

        # Atalho Ctrl + E (ou Ctrl + G) para pesquisar o grupo TIME BI OPTIMUS no Teams Desktop
        pyautogui.hotkey('ctrl', 'e')
        time.sleep(1.0)
        pyperclip.copy(grupo)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(2.0)
        pyautogui.press('down')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(2.5)

        # Atalho oficial do Teams para focar na caixa de mensagem
        pyautogui.hotkey('alt', 'shift', 'c')
        time.sleep(0.8)

        # Cola o resumo do monitoramento e envia
        pyperclip.copy(texto)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1.0)
        pyautogui.press('enter')
        time.sleep(1.0)

        log.info('Resumo enviado com sucesso no grupo %s do Teams Desktop!', grupo)
        return True
    except Exception as e:
        log.error('Falha ao enviar mensagem no Teams Desktop: %s', str(e)[:250])
        return False


def rotina_teams(cfg, res, forcar=False):
    t_cfg = cfg.get('teams', {})
    if not t_cfg.get('ativo', True) and not forcar:
        return

    est = json.loads(ESTADO.read_text(encoding='utf-8')) if ESTADO.exists() else {}
    agora = datetime.now()
    ult_str = est.get('ultimo_teams')
    ult_dt = datetime.fromisoformat(ult_str) if ult_str else None
    int_min = int(t_cfg.get('intervalo_min', 120))

    if forcar or ult_dt is None or (agora - ult_dt) >= timedelta(minutes=int_min):
        texto_teams = resumo(res, titulo='Monitoramento OptimusBI · TIME BI OPTIMUS')
        if enviar_teams_desktop(cfg, texto_teams):
            est['ultimo_teams'] = agora.isoformat()
            ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')


def gerar_pdf_visao_geral(cfg, res):
    """
    Gera um arquivo PDF executivo completo renderizando o painel atualizado (Visão Geral + Detalhamento)
    via Chrome Headless (Page.printToPDF), sem depender de bibliotecas externas pesadas.
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options

        html = AQUI / cfg.get('dashboard_html', 'dashboard.html')
        if not html.exists():
            return None

        pasta_destino = Path(cfg.get('pasta_dashboard_acompanhamento', DEFAULT['pasta_dashboard_acompanhamento']))
        try:
            pasta_destino.mkdir(parents=True, exist_ok=True)
        except Exception:
            pasta_destino = AQUI

        arq_pdf = pasta_destino / f"Visao_Geral_OptimusBI_{datetime.now():%Y-%m-%d_%Hh%M}.pdf"

        op = Options()
        for a in ('--headless=new', '--disable-gpu', '--window-size=1440,2200', f"--user-data-dir={AQUI / 'chrome_pdf_tmp'}"):
            op.add_argument(a)
        exe = _achar_chrome(cfg)
        if exe:
            op.binary_location = exe

        drv = webdriver.Chrome(options=op)
        try:
            drv.get(html.as_uri())
            time.sleep(5)
            # Exibe todas as seções para o PDF ficar com a visão geral + tabelas completas
            drv.execute_script("""
                var t = document.getElementById('toasts'); if(t) t.style.display='none';
                document.querySelectorAll('main > section').forEach(s => s.classList.remove('hidden'));
            """)
            time.sleep(1.5)
            pdf_data = drv.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True,
                "landscape": True,
                "paperWidth": 11.69,
                "paperHeight": 8.27,
                "marginTop": 0.3,
                "marginBottom": 0.3,
                "marginLeft": 0.3,
                "marginRight": 0.3
            })
            arq_pdf.write_bytes(base64.b64decode(pdf_data['data']))
            log.info('PDF Visão Geral gerado com sucesso em: %s', arq_pdf)
            return arq_pdf
        finally:
            drv.quit()
    except Exception as e:
        log.error('Falha ao gerar PDF da Visão Geral: %s', str(e)[:250])
        return None


def montar_corpo_email_explicativo(res):
    """Monta o texto explicativo completo para o corpo do e-mail (07h e 17h)."""
    ativos = [x for x in res if x['status'] != 'nd']
    ok_list = [x for x in ativos if x['status'] == 'ok']
    warn_list = [x for x in ativos if x['status'] == 'warn']
    err_list = [x for x in ativos if x['status'] == 'err']
    bis = [x for x in ativos if x.get('categoria') == 'pbix']
    saude = round(100 * len(ok_list) / len(ativos)) if ativos else 0

    linhas = [
        f"Olá, equipe TIME BI OPTIMUS!",
        f"",
        f"Segue a Visão Geral do Monitoramento OptimusBI referente à leitura de {datetime.now():%d/%m/%Y às %H:%M}.",
        f"O relatório visual completo encontra-se em PDF anexo a este e-mail.",
        f"",
        f"==================================================================",
        f"1. RESUMO EXECUTIVO DO ACOMPANHAMENTO",
        f"==================================================================",
        f"• Saúde Geral do Ambiente : {saude}%",
        f"• Total de Itens Ativos   : {len(ativos)} processos/bases/painéis",
        f"• Em Dia (OK)             : {len(ok_list)}",
        f"• Em Alerta (Atenção)     : {len(warn_list)}",
        f"• Críticos / Com Falha    : {len(err_list)}",
        f"• Painéis BI (.pbix) OK   : {sum(1 for x in bis if x['status'] == 'ok')} de {len(bis)}",
        f""
    ]

    if err_list:
        linhas += [
            f"==================================================================",
            f"2. DETALHAMENTO DE FALHAS / ATRASOS CRÍTICOS ({len(err_list)})",
            f"=================================================================="
        ]
        for x in err_list:
            motivo = "; ".join(x.get('erros', [])) or "Verificar painel"
            linhas.append(f"✖ [{x['prio']}] {x['nome']} (Resp: {x.get('responsavel') or 'Equipe'})")
            linhas.append(f"   Explicação: {motivo}")
        linhas.append("")

    if warn_list:
        linhas += [
            f"==================================================================",
            f"3. PONTOS DE ATENÇÃO / ALERTAS ({len(warn_list)})",
            f"=================================================================="
        ]
        for x in warn_list:
            aviso = "; ".join(x.get('avisos', [])) or "Alerta de janela/histórico"
            linhas.append(f"⚠ [{x['prio']}] {x['nome']} (Resp: {x.get('responsavel') or 'Equipe'}) -> {aviso}")
        linhas.append("")

    linhas += [
        f"==================================================================",
        f"4. PROCESSOS E PAINÉIS EM DIA ({len(ok_list)})",
        f"==================================================================",
        ", ".join(x['nome'].replace('PBIX · ', '') for x in ok_list),
        f"",
        f"Atenciosamente,",
        f"Rotina Automática - Monitoramento OptimusBI"
    ]
    return "\n".join(linhas)


def enviar_email_outlook_navegador(cfg, res):
    """
    Abre o Outlook Web no navegador (usando o perfil logado), cria um novo e-mail,
    preenche destinatários, assunto, corpo explicativo e anexa o PDF da Visão Geral.
    """
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    e_cfg = cfg.get('email_outlook', DEFAULT['email_outlook'])
    porta = int(e_cfg.get('porta_chrome', 9226))
    perfil = e_cfg.get('perfil_chrome', r'C:\chrome_outlook_planner')
    destinatarios = e_cfg.get('destinatarios', [])
    if not destinatarios:
        log.warning('Nenhum destinatário configurado em email_outlook.destinatarios.')
        return False

    arq_pdf = gerar_pdf_visao_geral(cfg, res)
    corpo_texto = montar_corpo_email_explicativo(res)
    assunto = f"{e_cfg.get('assunto_prefixo', '[OptimusBI] Visão Geral')} - {datetime.now():%d/%m/%Y %H:%M}"

    proc = None
    try:
        log.info('Iniciando disparo de e-mail via Outlook no navegador...')
        encerrar_chrome_porta(porta)
        exe = _achar_chrome(cfg)
        if not exe:
            log.error('chrome.exe não encontrado para o Outlook Web.')
            return False

        proc = subprocess.Popen([
            exe,
            f'--remote-debugging-port={porta}',
            f'--user-data-dir={perfil}',
            '--no-first-run',
            'https://outlook.office.com/mail/deeplink/compose'
        ])

        for _ in range(35):
            if porta_aberta(porta):
                time.sleep(4)
                break
            time.sleep(1)

        op = Options()
        op.debugger_address = f"127.0.0.1:{porta}"
        drv = webdriver.Chrome(options=op)
        drv.get('https://outlook.office.com/mail/deeplink/compose')
        time.sleep(8)

        # 1. Preenche o campo "Para"
        campo_para = WebDriverWait(drv, 45).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//div[@role='textbox' and (@aria-label='Para' or @aria-label='To')] | //input[contains(@aria-label,'Para') or contains(@aria-label,'To')]"
            ))
        )
        campo_para.click()
        for dest in destinatarios:
            campo_para.send_keys(dest)
            time.sleep(0.8)
            campo_para.send_keys(Keys.ENTER)
            time.sleep(0.5)

        # 2. Preenche o Assunto
        campo_assunto = WebDriverWait(drv, 20).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//input[contains(@placeholder,'assunto') or contains(@placeholder,'subject') or contains(@aria-label,'Assunto') or contains(@aria-label,'Subject')]"
            ))
        )
        campo_assunto.click()
        campo_assunto.send_keys(assunto)
        time.sleep(1)

        # 3. Preenche o Corpo Explicativo
        campo_corpo = WebDriverWait(drv, 20).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//div[@role='textbox' and (contains(@aria-label,'Corpo da mensagem') or contains(@aria-label,'Message body'))]"
            ))
        )
        campo_corpo.click()
        campo_corpo.send_keys(corpo_texto)
        time.sleep(1.5)

        # 4. Anexa o PDF da Visão Geral
        if arq_pdf and arq_pdf.exists():
            inputs_file = drv.find_elements(By.XPATH, "//input[@type='file']")
            if inputs_file:
                inp_anexo = inputs_file[0]
                drv.execute_script("arguments[0].style.display = 'block';", inp_anexo)
                inp_anexo.send_keys(str(arq_pdf.resolve()))
                time.sleep(6)

        # 5. Envia o e-mail (Ctrl + Enter no Outlook Web ou clique no botão Enviar)
        campo_corpo.send_keys(Keys.CONTROL, Keys.ENTER)
        time.sleep(2)
        botoes_enviar = drv.find_elements(By.XPATH, "//button[@aria-label='Enviar' or @aria-label='Send']")
        for b in botoes_enviar:
            if b.is_displayed():
                try:
                    b.click()
                    break
                except Exception:
                    pass

        time.sleep(6)
        log.info('E-mail com PDF enviado com sucesso pelo Outlook Web!')
        return True

    except Exception as e:
        log.error('Falha ao enviar e-mail no Outlook Web: %s', str(e)[:250])
        return False
    finally:
        if proc:
            try:
                proc.terminate()
            except Exception:
                pass
        encerrar_chrome_porta(porta)


def dar_check_planner_navegador(cfg):
    """
    Ao final do dia, abre o Microsoft Planner no navegador e dá check (conclui)
    nas atividades pendentes atribuídas a você.
    """
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By

    p_cfg = cfg.get('planner', DEFAULT['planner'])
    porta = int(p_cfg.get('porta_chrome', 9226))
    perfil = p_cfg.get('perfil_chrome', r'C:\chrome_outlook_planner')
    url_planner = p_cfg.get('url_planner', 'https://planner.cloud.microsoft/')
    palavras = [k.lower() for k in p_cfg.get('palavras_chave_tarefas', [])]

    proc = None
    try:
        log.info('Iniciando rotina de check nas atividades do Planner (%s)...', url_planner)
        encerrar_chrome_porta(porta)
        exe = _achar_chrome(cfg)
        if not exe:
            log.error('chrome.exe não encontrado para o Planner.')
            return False

        proc = subprocess.Popen([
            exe,
            f'--remote-debugging-port={porta}',
            f'--user-data-dir={perfil}',
            '--no-first-run',
            url_planner
        ])

        for _ in range(35):
            if porta_aberta(porta):
                time.sleep(4)
                break
            time.sleep(1)

        op = Options()
        op.debugger_address = f"127.0.0.1:{porta}"
        drv = webdriver.Chrome(options=op)
        drv.get(url_planner)
        time.sleep(10)

        # Localiza os botões de conclusão (check) das tarefas não concluídas no Planner
        seletores_check = (
            "//button[contains(@aria-label,'Concluir tarefa') or contains(@aria-label,'Complete task') "
            "or contains(@aria-label,'Marcar como concluída') or contains(@title,'Concluir')]"
            " | //div[@role='checkbox' and @aria-checked='false']"
        )
        botoes = drv.find_elements(By.XPATH, seletores_check)
        concluidas = 0

        for btn in botoes:
            try:
                if not btn.is_displayed():
                    continue
                texto_card = (btn.get_attribute('aria-label') or '') + ' ' + (btn.find_element(By.XPATH, "./ancestor::*[3]").text or '')
                if palavras and not any(p in texto_card.lower() for p in palavras):
                    continue
                btn.click()
                concluidas += 1
                time.sleep(1.2)
            except Exception:
                continue

        log.info('Rotina do Planner finalizada: %s atividade(s) marcada(s) com check!', concluidas)
        time.sleep(3)
        return True

    except Exception as e:
        log.error('Falha ao dar check no Planner: %s', str(e)[:250])
        return False
    finally:
        if proc:
            try:
                proc.terminate()
            except Exception:
                pass
        encerrar_chrome_porta(porta)


def rotina_email_e_planner_diario(cfg, res, forcar_email=False, forcar_planner=False):
    """
    Garante o disparo do e-mail às 07:00 e às 17:00 e o check no Planner ao final do dia (17:00+).
    """
    est = json.loads(ESTADO.read_text(encoding='utf-8')) if ESTADO.exists() else {}
    agora = datetime.now()
    hoje_iso = date.today().isoformat()
    hm_atual = agora.strftime('%H:%M')

    # 1. Rotina de E-mail (07:00 e 17:00)
    e_cfg = cfg.get('email_outlook', DEFAULT['email_outlook'])
    if e_cfg.get('ativo', True) or forcar_email:
        enviados_hoje = est.get('emails_enviados_dia', {})
        if enviados_hoje.get('dia') != hoje_iso:
            enviados_hoje = {'dia': hoje_iso, 'janelas': []}

        for h_alvo in e_cfg.get('horarios', ['07:00', '17:00']):
            if forcar_email or (hm_atual >= h_alvo and h_alvo not in enviados_hoje['janelas']):
                if enviar_email_outlook_navegador(cfg, res):
                    if h_alvo not in enviados_hoje['janelas']:
                        enviados_hoje['janelas'].append(h_alvo)
                    est['emails_enviados_dia'] = enviados_hoje
                    ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')
                if forcar_email:
                    break

    # 2. Rotina do Planner ao final do dia (a partir das 17:05)
    p_cfg = cfg.get('planner', DEFAULT['planner'])
    if p_cfg.get('ativo', True) or forcar_planner:
        h_planner = p_cfg.get('horario_check', '17:05')
        ja_fez_planner = (est.get('ultimo_planner_check') == hoje_iso)
        if forcar_planner or (hm_atual >= h_planner and not ja_fez_planner):
            if dar_check_planner_navegador(cfg):
                est['ultimo_planner_check'] = hoje_iso
                ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')


# ==============================================================================
# 8. TEMPLATE DO DASHBOARD LOCAL (HTML v6.0 - COM BOTÕES TEAMS, PDF, OUTLOOK E PLANNER)
# ==============================================================================
VERSAO_PAINEL = 'v6.0 · Teams | PDF | Outlook 07h/17h | Planner'

HTML = r'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
  <title>Monitoramento OptimusBI v6.0</title>
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
    .bar { display: flex; gap: 6px; flex-wrap: wrap; align-items: center; }
    button, select { font: inherit; font-size: 12.5px; font-weight: 600; border: 1px solid var(--slate-200); background: var(--card); color: var(--ink); padding: 7px 12px; border-radius: 6px; cursor: pointer; transition: 0.15s; }
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
    .k b { font-size: 32px; display: block; line-height: 1.1; color: var(--ink); }
    .k span { color: var(--slate-600); font-size: 12px; }
    .box { background: var(--card); border: 1px solid var(--slate-200); border-radius: 8px; padding: 14px; }
    .box h3 { margin: 0 0 10px; font-size: 13.5px; color: var(--blue-inst); font-weight: 700; }
    .box.big { padding: 18px; } .box.big h3 { font-size: 15px; }
    .cw { position: relative; height: 380px; }
    .charts2 { display: grid; grid-template-columns: repeat(auto-fit, minmax(560px, 1fr)); gap: 16px; margin-bottom: 18px; }
    .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(420px, 1fr)); gap: 14px; }
    .grid1 { display: block; }
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
    .lt { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 11.5px; }
    .lt td, .lt th { border-bottom: 1px solid var(--slate-200); padding: 3px 4px; text-align: left; }
    .lt .s-err { color: var(--red); font-weight: 700; } .lt .s-warn { color: var(--amber-800); font-weight: 600; } .lt .s-ok { color: var(--blue-inst); }
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
    @media(max-width:600px){ #toasts { width: 240px; } .charts2, .grid { grid-template-columns: 1fr; } }
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
      <button class="p" id="bRefresh">🔄 Atualizar</button>
      <button class="g" id="bTeams" title="Enviar agora no grupo TIME BI OPTIMUS do Teams Desktop">👥 Enviar p/ Teams</button>
      <button class="g" id="bEmail" title="Gerar PDF e disparar por e-mail no Outlook Web">📧 Enviar PDF Outlook</button>
      <button class="g" id="bPlanner" title="Entrar no Planner e dar check nas minhas atividades">✅ Check Planner</button>
      <button class="g" id="bVercel">☁️ Vercel (.xlsx)</button>
      <button class="g" id="bWa">💬 WhatsApp</button>
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
      <strong>Teams:</strong> TIME BI OPTIMUS | <strong>E-mail + PDF:</strong> 07:00 e 17:00 | <strong>Planner Check:</strong> 17:05 | <strong>Vercel:</strong> 30m
    </div>
  </div>

  <section id="s-geral">
    <div class="sum" id="sum"></div>
    <div class="box big" style="margin-bottom:16px"><h3>Monitoramento dos BIs — horas desde a última atualização (cor = situação, nome = BI · responsável)</h3><div class="cw" id="cwBis"><canvas id="cBis"></canvas></div></div>
    <div class="box big" style="margin-bottom:16px">
      <h3>Evolução da equipe OptimusBI — % de itens em dia por responsável</h3>
      <div id="team"></div>
    </div>
    <div class="charts2">
      <div class="box big"><h3>Evolução por hora (hoje) — % em dia (colunas) e falhas (linha)</h3><div class="cw"><canvas id="cHoraE"></canvas></div></div>
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
  <div class="dh"><b>Resumo para WhatsApp / Teams</b><button id="bWc">Fechar</button></div>
  <div class="db">
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
const PESSOA = { Karine: '#F24405', Agenor: '#F2B705', Miller: '#122A44', Daniel: '#6B7280' };
const resp = r => r.responsavel || 'Equipe OptimusBI';
const donos = r => (r.responsavel || '').split(/\s+e\s+|,/).map(x => x.trim()).filter(Boolean);
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
  $('#bRefresh').disabled = true; $('#bRefresh').textContent = '⏳...';
  try { await fetch('/api/atualizar'); } catch (e) {}
  location.reload();
};
$('#bTeams').onclick = async () => {
  $('#bTeams').disabled = true; $('#bTeams').textContent = '⏳ Teams...';
  try { await fetch('/api/teams'); alert('Resumo enviado para o grupo TIME BI OPTIMUS no Teams Desktop!'); }
  catch (e) { alert('Erro ao acionar o Teams.'); }
  $('#bTeams').disabled = false; $('#bTeams').textContent = '👥 Enviar p/ Teams';
};
$('#bEmail').onclick = async () => {
  $('#bEmail').disabled = true; $('#bEmail').textContent = '⏳ PDF + E-mail...';
  try { await fetch('/api/email'); alert('PDF gerado e enviado por e-mail no Outlook Web!'); }
  catch (e) { alert('Erro ao enviar e-mail.'); }
  $('#bEmail').disabled = false; $('#bEmail').textContent = '📧 Enviar PDF Outlook';
};
$('#bPlanner').onclick = async () => {
  $('#bPlanner').disabled = true; $('#bPlanner').textContent = '⏳ Planner...';
  try { await fetch('/api/planner'); alert('Rotina de check no Planner executada!'); }
  catch (e) { alert('Erro ao acessar o Planner.'); }
  $('#bPlanner').disabled = false; $('#bPlanner').textContent = '✅ Check Planner';
};
$('#bVercel').onclick = async () => {
  $('#bVercel').disabled = true; $('#bVercel').textContent = '⏳ Vercel...';
  try { const j = await (await fetch('/api/vercel')).json(); alert('Enviado p/ Vercel!\nArquivo: ' + j.arquivo); }
  catch (e) {}
  $('#bVercel').disabled = false; $('#bVercel').textContent = '☁️ Vercel (.xlsx)';
};

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

function render() {
  const hVal = $('#hourSel').value;
  let list = D.r;
  if (hVal !== 'now' && snaps[hVal]) {
    list = snaps[hVal].itens;
    $('#snapInfo').innerHTML = `<b style="color:var(--orange)">Exibindo fotografia das ${hVal} (salva às ${snaps[hVal].salvo_as})</b>`;
  } else $('#snapInfo').textContent = 'Pastas lidas automaticamente pelo Python.';
  $$('main > section').forEach(s => s.classList.toggle('hidden', s.id !== 's-' + tab));
  const ativos = list.filter(r => r.status !== 'nd');
  if (tab === 'geral') renderGeral(list, ativos);
  if (tab === 'pbix') renderBis(list);
  if (tab === 'logs') renderLogs(list);
  if (tab === 'det') renderDet(list);
  renderToasts(hVal);
}

function renderGeral(list, ativos) {
  const c = { ok: 0, warn: 0, err: 0 };
  ativos.forEach(r => { if (c[r.status] != null) c[r.status]++; });
  const bis = ativos.filter(r => r.categoria === 'pbix'), bOk = bis.filter(r => r.status === 'ok').length;
  const saude = ativos.length ? Math.round(100 * c.ok / ativos.length) : 0;
  $('#sum').innerHTML = [
    ['Itens monitorados', ativos.length, 'var(--blue-inst)'], ['Em dia (OK)', c.ok, 'var(--blue-inst)'],
    ['Alertas (Média)', c.warn, 'var(--gold)'], ['Crítico / Erro', c.err, 'var(--red)'],
    ['BIs em dia', `${bOk}/${bis.length}`, 'var(--blue-dark)'], ['Saúde geral', saude + '%', saude >= 90 ? 'var(--blue-inst)' : saude >= 70 ? 'var(--gold)' : 'var(--red)']
  ].map(([a, b, col]) => `<div class="k" style="--c:${col}"><b>${b}</b><span>${a}</span></div>`).join('');

  const px = ordena(list.filter(r => r.categoria === 'pbix' && r.pbix && r.pbix.arquivo));
  $('#cwBis').style.height = Math.max(380, px.length * 30 + 60) + 'px';
  mk('cBis', {
    type: 'bar',
    data: { labels: px.map(r => `${r.nome.replace('PBIX · ', '')} · ${resp(r)}`), datasets: [{ data: px.map(r => r.pbix.idade_h), backgroundColor: px.map(r => COR[r.status]), borderRadius: 4 }] },
    options: { indexAxis: 'y', maintainAspectRatio: false, animation: false, plugins: { legend: { display: false } }, scales: { x: { title: { display: true, text: 'horas' } } } }
  });

  const eq = {};
  const add = (n, r) => { const e = eq[n] || (eq[n] = [0, 0]); e[1]++; if (r.status === 'ok') e[0]++; };
  ativos.forEach(r => donos(r).forEach(n => add(n, r)));
  const linhas = [['Equipe OptimusBI (geral)', ativos.filter(r => r.status === 'ok').length, ativos.length], ...Object.entries(eq).sort().map(([n, v]) => [n, v[0], v[1]])];
  $('#team').innerHTML = linhas.map(([n, ok, t], i) => {
    const pct = t ? Math.round(100 * ok / t) : 0, cor = pct >= 90 ? 'var(--blue-inst)' : pct >= 70 ? 'var(--gold)' : 'var(--red)';
    return `<div class="pb ${i === 0 ? 'pb0' : ''}"><div class="pb-h"><b>${i ? '👤 ' : '👥 '}${n}</b><span>${ok} de ${t} em dia · <b>${pct}%</b></span></div><div class="pb-t"><div class="pb-f" data-w="${pct}" style="background:${cor}"></div></div></div>`;
  }).join('');
  requestAnimationFrame(() => setTimeout(() => $$('.pb-f').forEach(el => el.style.width = el.dataset.w + '%'), 60));

  const hoje = new Date().toISOString().slice(0, 10), H = D.hist || [];
  const pct = x => { const t = x.ok + x.warn + x.err; return t ? Math.round(100 * x.ok / t) : 0; };
  const porHora = {}; H.filter(x => x.dia === hoje).forEach(x => porHora[x.h.slice(0, 2)] = x);
  const hk = Object.keys(porHora).sort();
  mk('cHoraE', {
    data: { labels: hk.map(h => h + 'h'), datasets: [
      { type: 'bar', label: '% em dia', data: hk.map(h => pct(porHora[h])), backgroundColor: '#0339A6', borderRadius: 4, yAxisID: 'y' },
      { type: 'line', label: 'Falhas', data: hk.map(h => porHora[h].err), borderColor: '#F21D2F', backgroundColor: '#F21D2F', tension: .3, yAxisID: 'y1' }] },
    options: { maintainAspectRatio: false, animation: false, scales: { y: { min: 0, max: 100 }, y1: { position: 'right', grid: { drawOnChartArea: false }, beginAtZero: true } } }
  });
  const porDia = {}; H.forEach(x => porDia[x.dia] = x);
  const dk = Object.keys(porDia).sort().slice(-20);
  const nomes = [...new Set(dk.flatMap(d => Object.keys(porDia[d].eq || {})))].sort();
  mk('cDiaE', {
    type: 'line',
    data: { labels: dk.map(d => d.slice(8) + '/' + d.slice(5, 7)), datasets: [
      { label: 'Equipe (geral)', data: dk.map(d => pct(porDia[d])), borderColor: '#0339A6', backgroundColor: '#0339A6', borderWidth: 4, tension: .3 },
      ...nomes.map(n => ({ label: n, data: dk.map(d => { const e = (porDia[d].eq || {})[n]; return e && e[1] ? Math.round(100 * e[0] / e[1]) : null; }), borderColor: PESSOA[n] || '#94A3B8', tension: .3, borderDash: [5, 4], spanGaps: true }))] },
    options: { maintainAspectRatio: false, animation: false, scales: { y: { min: 0, max: 100 } } }
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
  $('#gridLogs').innerHTML = `<div class="box"><h3>Auditoria da Pasta 52</h3>${(it.erros || []).map(x => `<div class="msg">${x}</div>`).join('')}${logsTable(ls, true)}</div>`;
}

function renderDet(list) {
  const base = list.filter(r => r.categoria !== 'logs');
  const tipo = r => r.categoria === 'pbix' ? 'pbix' : (r.tipo_proc || 'outros');
  const cn = k => base.filter(r => tipo(r) === k).length;
  chips('#chipsDet', [['todos', 'Todos', base.length], ['fluxo', '⚙️ Fluxo de execução', cn('fluxo')], ['download', '⬇️ Download de bases', cn('download')], ['pbix', '📊 BIs (.pbix)', cn('pbix')], ['outros', 'Outros', cn('outros')]], fDet, k => fDet = k);
  $('#gridDet').innerHTML = ordena(base.filter(r => fDet === 'todos' || tipo(r) === fDet)).map(card).join('') || '<div class="msg i">Nada neste filtro.</div>';
}

function renderToasts(hVal) {
  const al = ordena(D.r.filter(r => r.status === 'err'));
  const dis = JSON.parse(localStorage.getItem('optimusDismiss') || '{}'), now = Date.now();
  const key = r => r.nome + '|' + (r.erros[0] || '');
  const vis = al.filter(r => !(dis[key(r)] && now - dis[key(r)] < 30 * 60000));
  $('#bellN').textContent = al.length;
  if (hVal !== 'now') { $('#toasts').innerHTML = ''; return; }
  const cor = r => r.prio === 'P0' ? 'var(--red)' : r.prio === 'P1' ? 'var(--orange)' : 'var(--gold)';
  $('#toasts').innerHTML = vis.slice(0, 5).map(r => {
    const e = (r.erros[0] || 'Falha detectada');
    return `<div class="toast" style="--c:${cor(r)}" data-k="${encodeURIComponent(key(r))}"><button class="x">×</button>
      <div class="t1">🚨 [${r.prio}] ${r.nome.replace('PBIX · ', '')}</div><div class="t2">${e.length > 140 ? e.slice(0, 140) + '…' : e}</div>
      <div class="t3">👤 Resp: <b>${resp(r)}</b></div></div>`;
  }).join('');
  $$('#toasts .x').forEach(b => b.onclick = () => { const k = decodeURIComponent(b.parentElement.dataset.k); dis[k] = Date.now(); localStorage.setItem('optimusDismiss', JSON.stringify(dis)); render(); });
}
$('#bBell').onclick = () => { localStorage.removeItem('optimusDismiss'); render(); };

$('#last').textContent = 'Último monitoramento: ' + D.gerado;
setInterval(() => { $('#liveClock').textContent = fmtHMS(new Date()); secLeft = Math.max(0, secLeft - 1); $('#cd').textContent = secLeft + 's'; if (secLeft === 0 && !$('#dWa').open) location.reload(); }, 1000);
$('#bWa').onclick = () => {
  $('#waTxt').value = D.texto;
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
    ativos = [x for x in res if x['status'] != 'nd']
    c = lambda s: sum(x['status'] == s for x in ativos)
    agora_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    eq = {}
    for x in ativos:
        for p in [q.strip() for q in re.split(r'\s+e\s+|,', x.get('responsavel') or '') if q.strip()]:
            e = eq.setdefault(p, [0, 0]); e[1] += 1; e[0] += (x['status'] == 'ok')
    hist.append(dict(dia=date.today().isoformat(), eq=eq, h=datetime.now().strftime('%H:%M:%S'), ok=c('ok'), warn=c('warn'), err=c('err'),
                     bi_ok=sum(1 for x in ativos if x.get('categoria') == 'pbix' and x['status'] == 'ok'),
                     bi_err=sum(1 for x in ativos if x.get('categoria') == 'pbix' and x['status'] == 'err')))
    hist = hist[-500:]
    HIST.write_text(json.dumps(hist, ensure_ascii=False), encoding='utf-8')

    snaps_hoje = atualizar_snapshots_hora(res)
    dados = dict(
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
# 9. WHATSAPP A CADA 6 HORAS
# ==============================================================================
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
    return m


def resumo(res, titulo='Monitoramento OptimusBI', max_chars=950):
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


def _achar_chrome(cfg):
    cands = [cfg['whatsapp'].get('chrome_exe', ''), r'C:\Program Files\Google\Chrome\Application\chrome.exe',
             r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe')]
    return next((c for c in cands if c and os.path.exists(c)), None)


def capturar_visao_geral(cfg):
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        html = AQUI / cfg.get('dashboard_html', 'dashboard.html')
        if not html.exists():
            return None
        op = Options()
        for a in ('--headless=new', '--hide-scrollbars', '--force-device-scale-factor=1', '--window-size=1500,1800', '--disable-gpu',
                  f"--user-data-dir={AQUI / 'chrome_print_tmp'}"):
            op.add_argument(a)
        exe = _achar_chrome(cfg)
        if exe:
            op.binary_location = exe
        drv = webdriver.Chrome(options=op)
        try:
            drv.get(html.as_uri())
            time.sleep(6)
            drv.execute_script("var t=document.getElementById('toasts'); if(t) t.style.display='none';")
            altura = min(3400, int(drv.execute_script('return document.documentElement.scrollHeight')) + 20)
            drv.set_window_size(1500, altura)
            time.sleep(2)
            img = AQUI / 'visao_geral.png'
            drv.save_screenshot(str(img))
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


def encerrar_chrome_porta(porta):
    if os.name == 'nt':
        cmd = f'powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object {{ $_.Name -eq \'chrome.exe\' -and $_.CommandLine -like \'*{porta}*\' }} | Invoke-CimMethod -MethodName Terminate" >nul 2>&1'
        subprocess.run(cmd, shell=True)
    time.sleep(2)


def abrir_chrome_limpo(cfg):
    w = cfg['whatsapp']
    p = w['porta_chrome']
    exe = _achar_chrome(cfg)
    if not exe:
        return None
    perfil = w.get('perfil_chrome') or r'C:\chrome_whatsapp_monitor'
    proc = subprocess.Popen([
        exe, f'--remote-debugging-port={p}', f'--user-data-dir={perfil}', '--no-first-run', 'https://web.whatsapp.com'
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
    caixa = drv.find_elements(By.XPATH, "//div[@contenteditable='true']")[-1]
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
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    w = cfg['whatsapp']
    porta = w['porta_chrome']
    destinos = w.get('destinos', [])
    ok_all = True
    xp_caixa = "//footer//div[@contenteditable='true']"

    for idx, d in enumerate(destinos, start=1):
        proc = None
        try:
            encerrar_chrome_porta(porta)
            proc = abrir_chrome_limpo(cfg)
            if not proc:
                ok_all = False
                continue
            op = Options()
            op.debugger_address = f"127.0.0.1:{porta}"
            drv = webdriver.Chrome(options=op)

            enviado = False
            if imagem and Path(imagem).exists():
                try:
                    drv.get(f"https://web.whatsapp.com/send?phone={d['numero']}")
                    WebDriverWait(drv, 90).until(EC.presence_of_element_located((By.XPATH, xp_caixa)))
                    time.sleep(2)
                    _enviar_imagem(drv, imagem, texto)
                    enviado = True
                except Exception:
                    pass
            if not enviado:
                drv.get(f"https://web.whatsapp.com/send?phone={d['numero']}&text={quote(texto)}")
                caixa = WebDriverWait(drv, 90).until(EC.presence_of_element_located((By.XPATH, xp_caixa)))
                time.sleep(2)
                caixa.send_keys(Keys.ENTER)
                time.sleep(5)
        except Exception as e:
            ok_all = False
            log.error('Falha WhatsApp %s: %s', d['nome'], str(e)[:200])
        finally:
            if proc:
                try:
                    proc.terminate()
                except Exception:
                    pass
            encerrar_chrome_porta(porta)

    return ok_all


def rotina_whatsapp(cfg, res, enviar):
    w = cfg['whatsapp']
    if not (enviar and w.get('ativo')):
        return
    est = json.loads(ESTADO.read_text(encoding='utf-8')) if ESTADO.exists() else {}
    agora = datetime.now()
    freq_min = int(w.get('resumo_a_cada_min', 360))
    ultimo = datetime.fromisoformat(est['ultimo_resumo']) if est.get('ultimo_resumo') else None
    ativos = [x for x in res if x['status'] != 'nd']
    erros = sorted(f"{x['nome']}|{e}" for x in ativos for e in x['erros'])
    sig = '\n'.join(erros)

    if ultimo is None or (agora - ultimo) >= timedelta(minutes=freq_min):
        titulo = 'ALERTA - Monitoramento OptimusBI' if erros else 'Monitoramento OptimusBI'
        if enviar_whatsapp(cfg, resumo(res, titulo), capturar_visao_geral(cfg)):
            est['ultimo_resumo'] = agora.isoformat()
            est['ultima_assinatura'] = sig
            ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')


# ==============================================================================
# 10. SERVIDOR WEB EMBUTIDO (ANTI-CACHE + LIBERAÇÃO DA PORTA 8080) E CICLO
# ==============================================================================
ULTIMO_ESTADO_GLOBAL = {'cfg': None, 'res': [], 'snaps': {}, 'hist': []}


def liberar_porta_8080(porta=8080):
    """Mata automaticamente qualquer processo python antigo que esteja prendendo a porta 8080."""
    if os.name == 'nt':
        meu_pid = os.getpid()
        cmd = (
            f'powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort {porta} -ErrorAction SilentlyContinue | '
            f'Select-Object -ExpandProperty OwningProcess -Unique | Where-Object {{ $_ -ne {meu_pid} }} | '
            f'ForEach-Object {{ Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }}" >nul 2>&1'
        )
        subprocess.run(cmd, shell=True)
        time.sleep(1.5)


class PainelHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        return

    def _headers_anti_cache(self, content_type='application/json; charset=utf-8'):
        self.send_response(200)
        self.send_header('Content-type', content_type)
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        self.end_headers()

    def do_GET(self):
        if self.path.startswith('/api/atualizar'):
            ciclo(enviar=False)
            self._headers_anti_cache()
            self.wfile.write(json.dumps({'ok': True, 'hora': datetime.now().strftime('%H:%M:%S')}).encode('utf-8'))
            return

        if self.path.startswith('/api/teams'):
            cfg = ULTIMO_ESTADO_GLOBAL['cfg'] or sincronizar_config_automatico()
            rotina_teams(cfg, ULTIMO_ESTADO_GLOBAL['res'], forcar=True)
            self._headers_anti_cache()
            self.wfile.write(json.dumps({'ok': True}).encode('utf-8'))
            return

        if self.path.startswith('/api/email'):
            cfg = ULTIMO_ESTADO_GLOBAL['cfg'] or sincronizar_config_automatico()
            rotina_email_e_planner_diario(cfg, ULTIMO_ESTADO_GLOBAL['res'], forcar_email=True)
            self._headers_anti_cache()
            self.wfile.write(json.dumps({'ok': True}).encode('utf-8'))
            return

        if self.path.startswith('/api/planner'):
            cfg = ULTIMO_ESTADO_GLOBAL['cfg'] or sincronizar_config_automatico()
            rotina_email_e_planner_diario(cfg, ULTIMO_ESTADO_GLOBAL['res'], forcar_planner=True)
            self._headers_anti_cache()
            self.wfile.write(json.dumps({'ok': True}).encode('utf-8'))
            return

        if self.path.startswith('/api/vercel'):
            cfg = ULTIMO_ESTADO_GLOBAL['cfg'] or sincronizar_config_automatico()
            arq = gerar_relatorio_xlsx(cfg, ULTIMO_ESTADO_GLOBAL['res'], ULTIMO_ESTADO_GLOBAL['snaps'], ULTIMO_ESTADO_GLOBAL['hist'])
            enviar_xlsx_para_painel_vercel(cfg, arq)
            self._headers_anti_cache()
            self.wfile.write(json.dumps({'ok': True, 'arquivo': str(arq)}).encode('utf-8'))
            return

        if self.path.startswith('/logo.png') and (AQUI / 'logo.png').exists():
            self._headers_anti_cache('image/png')
            self.wfile.write((AQUI / 'logo.png').read_bytes())
            return

        arq_html = AQUI / (ULTIMO_ESTADO_GLOBAL['cfg'] or DEFAULT).get('dashboard_html', 'dashboard.html')
        if not arq_html.exists():
            ciclo(enviar=False)
        self._headers_anti_cache('text/html; charset=utf-8')
        self.wfile.write(arq_html.read_bytes())


def iniciar_servidor_web(porta=8080):
    liberar_porta_8080(porta)
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


def ciclo(enviar, forcar_relatorio=False, forcar_vercel=False, forcar_teams=False, forcar_email=False, forcar_planner=False):
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
        rotina_teams(cfg, res, forcar=forcar_teams)
    except Exception:
        log.error('Falha na rotina Teams\n%s', traceback.format_exc())

    try:
        rotina_email_e_planner_diario(cfg, res, forcar_email=forcar_email, forcar_planner=forcar_planner)
    except Exception:
        log.error('Falha na rotina Outlook/Planner\n%s', traceback.format_exc())

    try:
        rotina_vercel_30_minutos(cfg, res, snaps_hoje, hist, forcar=forcar_vercel)
    except Exception:
        log.error('Falha na rotina Vercel 30m\n%s', traceback.format_exc())

    try:
        verificar_relatorio_final_do_dia(cfg, res, snaps_hoje, hist, forcar=forcar_relatorio)
    except Exception:
        log.error('Falha ao gerar relatório diário .xlsx\n%s', traceback.format_exc())

    try:
        rotina_whatsapp(cfg, res, enviar)
    except Exception:
        log.error('Falha na rotina WhatsApp\n%s', traceback.format_exc())

    return cfg


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--uma-vez', action='store_true')
    ap.add_argument('--sem-whatsapp', action='store_true')
    ap.add_argument('--testar-teams', action='store_true', help='Testa o envio imediato no grupo TIME BI OPTIMUS do Teams Desktop')
    ap.add_argument('--testar-email', action='store_true', help='Gera o PDF da Visão Geral e envia agora pelo Outlook Web')
    ap.add_argument('--testar-planner', action='store_true', help='Abre o Planner no navegador e dá check nas atividades')
    ap.add_argument('--enviar-vercel', action='store_true')
    a = ap.parse_args()

    if a.testar_teams or a.testar_email or a.testar_planner:
        ciclo(
            enviar=False,
            forcar_teams=a.testar_teams,
            forcar_email=a.testar_email,
            forcar_planner=a.testar_planner
        )
        print('Teste concluído! Verifique o monitor.log para detalhes.')
        sys.exit()

    cfg_ini = sincronizar_config_automatico()
    if not a.uma_vez and not a.enviar_vercel:
        porta_web = int(cfg_ini.get('porta_servidor_web', 8080))
        iniciar_servidor_web(porta_web)
        log.info('Monitoramento OptimusBI %s iniciado (PID %s)', VERSAO_PAINEL, os.getpid())
        if cfg_ini.get('abrir_navegador_ao_iniciar', True):
            webbrowser.open(f'http://localhost:{porta_web}?v={int(time.time())}')

    while True:
        ini = time.time()
        try:
            cfg = ciclo(not a.sem_whatsapp, forcar_vercel=a.enviar_vercel)
        except Exception:
            log.error('Erro no ciclo\n%s', traceback.format_exc())
            cfg = DEFAULT
        if a.uma_vez or a.enviar_vercel:
            break
        tempo_espera = max(10, int(cfg.get('intervalo_seg', 60)) - (time.time() - ini))
        time.sleep(tempo_espera)