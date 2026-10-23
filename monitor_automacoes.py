#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitoramento OptimusBI - Completo e Atualizado:
- Verifica pastas, bases, BIs (.pbix) e logs (Pasta 52) a cada 60s
- Inclui aba exclusiva de Monitoramento de BIs com 4 gráficos e Modal Pop-up com som opcional
- A cada 30 minutos: gera o .xlsx em C:\\Users\\agenor.sales\\Desktop\\dashboard acompanhamento
  e faz o upload automático em https://acompanhameto-fluxo-diario-de-taref.vercel.app/ na aba Monitoramento
- A cada 6 horas (360 min): envia resumo no WhatsApp (1 mensagem por contacto, fechando e reiniciando o navegador)
- Às 18h: guarda também o relatório diário final na pasta G:\\...
"""
import argparse
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
# 1. CAMINHOS BASE, FICHEIROS DE ESTADO E LOGGING
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
    """Função auxiliar para estruturar a configuração de cada pasta monitorada."""
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


def _bi(nome, busca, prio, regra='diaria', **kw):
    """Função auxiliar para estruturar cada painel PBIX monitorado na Pasta 3."""
    return dict(nome=nome, busca=busca, prio=prio, regra=regra, **kw)


B7 = '7 - BASES DE RELATORIOS\\'
MM = '\\{YYYY}\\{MM} - {MES}'
_APRR_H = ['06:00', '09:00', '11:00', '13:00', '15:00', '17:00']

# ==============================================================================
# 2. CONFIGURAÇÃO PADRÃO E REGRAS DO MAPEAMENTO DE PASTAS
# ==============================================================================
DEFAULT = {
    'versao_config': 6,
    'raiz': r'G:\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES',
    'raiz_rede_fallback': r'\\10.1.17.4\Usuarios\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES',
    'pasta_logs': r'1 - RELATORIOS\52 - LOGS PROJETOS',
    'pasta_pbix': r'3 - ATUALIZAÇÃO PBIX\PBIX OFICIAL',
    'hora_limite_manha': '11:00',  # Painéis da manhã atualizados até às 11h
    'bi_checar_fim_de_semana': False,
    'logs_ignorar': [],
    'pasta_relatorio_diario': r'G:\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES\7 - BASES DE RELATORIOS\RELATORIO PAINEL MONITORAMENTO DIARIO',
    'pasta_dashboard_acompanhamento': r'C:\Users\agenor.sales\Desktop\dashboard acompanhamento',
    'url_painel_vercel': 'https://acompanhameto-fluxo-diario-de-taref.vercel.app/',
    'intervalo_vercel_min': 30,  # Alimenta o painel online a cada 30 minutos
    'hora_relatorio_diario': 18,
    'intervalo_seg': 60,         # Atualiza o ecrã a cada 60 segundos
    'porta_servidor_web': 8080,
    'abrir_navegador_ao_iniciar': True,
    'dashboard_html': 'dashboard.html',
    'whatsapp': {
        'ativo': True,
        'porta_chrome': 9222,
        'resumo_a_cada_min': 360,  # Envia no WhatsApp apenas a cada 6 horas (360 minutos)
        'horarios_resumo': ['00:00', '06:00', '12:00', '18:00', '23:59'],
        'alerta_imediato_erro': False,
        'perfil_chrome': r'C:\chrome_whatsapp_monitor',
        'chrome_exe': '',
        'destinos': [
            {'nome': 'Agenor', 'numero': '558589768794'},
            {'nome': 'Miller', 'numero': '553199300997'}
        ]
    },
    # PASTA 3 - TODOS OS BIs (PBIX OFICIAL)
    'bis_pbix': [
        # P0 - Críticos (a cada 2 horas / diários manhã)
        _bi('Relatório Alerta de Vagas v2', ['alerta', 'vagas', 'v2'], 'P0', 'intervalo_horas', intervalo_h=2, excluir=['sub', 'tele']),
        _bi('Alerta Diário de Vagas_Sub', ['alerta', 'vagas', 'sub'], 'P0', 'intervalo_horas', intervalo_h=2),
        _bi('BI_Alerta_Vagas_TeleSaude', ['vagas', 'tele'], 'P0', 'intervalo_horas', intervalo_h=2),
        _bi('BI_NECESSIDADE_COMPRA', ['necessidade', 'compra'], 'P0', 'diaria', hora_limite='11:00'),

        # P1 - Alta Prioridade
        _bi('Painel Médicos Hapclinicas V2', ['painel', 'medicos'], 'P1', 'diaria', hora_limite='11:00'),
        _bi('Relatório Indicadores Rede Cred', ['indicadores', 'rede'], 'P1', 'diaria', hora_limite='11:00'),
        _bi('Relatório de Consultas Disponibilizadas v2', ['consultas', 'disponibilizadas'], 'P1', 'diaria', hora_limite='11:00'),
        _bi('Captacao_de_Agendas', ['captacao', 'agendas'], 'P1', 'ate_hora', hora_limite='17:00', tolerancia_min=30),
        _bi('Controle_Agendas', ['controle', 'agendas'], 'P1', 'ate_hora', hora_limite='17:00', tolerancia_min=30),
        _bi('Farol de Indice Consultas por Hora - ICH', ['farol'], 'P1', 'ate_hora', hora_limite='14:00', tolerancia_min=30),
        _bi('Relatório de Estoque e Abastecimento v2', ['estoque', 'abastecimento'], 'P1', 'diaria', hora_limite='11:00'),
        _bi('Acompanhamento de Pagamentos Prestadores', ['pagamentos', 'prestadores'], 'P1', 'ultima', limite_dias=2),

        # P2 - Média Prioridade e Janelas Específicas
        _bi('ESCALA ESPECIAL', ['escala', 'especial'], 'P2', 'diaria', hora_limite='11:00'),
        _bi('Relatório Atraso dos Médicos', ['atraso', 'medicos'], 'P2', 'diaria', hora_limite='11:00'),
        _bi('RELATORIO DE CONSULTAS EM TRANSFERENCIA', ['transferencia'], 'P2', 'diaria', hora_limite='11:00'),
        _bi('Relatório de Cancelamento v2', ['cancelamento'], 'P2', 'diaria', hora_limite='11:00'),
        _bi('Relatório de Acompanhamento de Mensageria', ['mensageria'], 'P2', 'diaria', hora_limite='11:00'),
        _bi('Monitoramento de Marcação', ['monitoramento', 'marcacao'], 'P2', 'ate_hora', hora_limite='06:00', tolerancia_min=60),
        _bi('Relatório de Cirurgias', ['cirurgias'], 'P2', 'diaria', hora_limite='11:00'),
        _bi('Relatorio de Tagueamento', ['tagueamento'], 'P2', 'horarios', horarios=['08:00', '11:30', '17:00'], tolerancia_min=60),
        _bi('Relatorio_SAC (3x ao dia / Mensal)', ['sac'], 'P2', 'horarios', horarios=['06:30', '11:30', '17:30'], tolerancia_min=60, excluir=['nip']),
        _bi('Relatório de SAC e NIP (desenvolvimento)', ['sac', 'nip'], 'P2', 'horarios', horarios=['07:30', '09:00', '13:00', '22:30'], tolerancia_min=60),
        _bi('Acompanhamento de Valores Pagos por Especialidade', ['valores', 'pagos', 'especialidade'], 'P2', 'ate_hora', hora_limite='05:00', tolerancia_min=60),
        _bi('Relatório Check-up da mulher', ['check', 'mulher'], 'P2', 'ultima', limite_dias=3),
        _bi('Relatório_Forecast_Monitoramento_Limites_e_Mediana', ['forecast'], 'P2', 'ultima', limite_dias=3),
        _bi('Relatório Pacotes (3x: 08:30, 11:00, 17:30)', ['pacotes'], 'P2', 'ignorar', horarios=['08:30', '11:00', '17:30'])
    ],
    # MONITORES DE FLUXOS (PASTA 52) E BASES (PASTA 7 / PASTA 1)
    'monitores': [
        # 🔴 P0 — Crítica
        {
            'nome': 'BI Alerta Vagas v2', 'categoria': 'bi', 'prio': 'P0', 'regra': 'aprr',
            'log': 'log_execucao_alerta_de_vagas_{YYYY}_{MM}.xlsx', 'min_linhas': 40,
            'pastas': [
                F(n, B7 + n + ' - DIARIA' + MM + '\\{DD}', 2, '', True, 'aprr', False)
                for n in ('T22APRR', 'T22APRR_PROC', 'T22APRR_VS', 'T22APRR_TRANSF')
            ] + [
                F(
                    rotulo='PARQUET DA HORA (F_vagas_atual)',
                    caminho=r'1 - RELATORIOS\47 - BI_TELESSAUDE\FATO',
                    caminho_fallback=r'\\10.1.17.4\Usuarios\Credenciamento Medico\NUCLEO DE AGENDAS\RAFAEL FERNANDES\1 - RELATORIOS\47 - BI_TELESSAUDE\FATO',
                    esperado=1,
                    contem='F_vagas_atual.parquet',
                    so_hoje=True,
                    arquivos_obrigatorios=['F_vagas_atual.parquet']
                )
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

        # 🟠 P1 — Alta (Bases e Fluxos)
        {
            'nome': 'T0012 - Disponibilidade / Marcada', 'categoria': 'base', 'prio': 'P1', 'regra': '',
            'log': '', 'min_linhas': 0,
            'pastas': [
                F('T0012 DISPONIBILIDADE', B7 + r'T0012 - ULTIMA BASE DISPONIBILIDADE MENSAL\{YYYY}', 14, '', True, '', conf_mensal=True),
                F('T0012 MARCADA', B7 + r'T0012 - ULTIMA BASE MARCADA MENSAL\{YYYY}', 14, '', True, '', conf_mensal=True)
            ]
        },
        {
            'nome': 'Tabela DB mensal T0012', 'categoria': 'base', 'prio': 'P1', 'regra': '',
            'log': '', 'min_linhas': 0,
            'pastas': [
                F('DB_T0012 FLUXO PYTHON', B7 + r'DB_T0012\FLUXO PYTHON', 3, '', True, '', conf_mensal=True)
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
                F('T22J3', B7 + 'T22J3 - BASE DIARIA' + MM, 1, 'T22J3_{DD}{MM}{YYYY}', True, '', conf_mensal=True)
            ]
        },
        {
            'nome': 'T9016 - Base mensal (Dias 1, 5 e 10)', 'categoria': 'base', 'prio': 'P1', 'regra': '',
            'log': '', 'min_linhas': 0,
            'pastas': [
                F('T9016', B7 + 'T9016 - BASE MENSAL' + MM, 5, '', True, '', conf_mensal=False,
                  dias_mes=[1, 5, 10], prefixos=['T9016_', 'NDISPRJ', 'NDIMINAS', 'CLINIPAM', 'CCG'])
            ]
        },
        {
            'nome': 'Estoque e Abastecimento', 'categoria': 'bi', 'prio': 'P1', 'regra': '',
            'log': 'log_execucao_estoque_abastecimento_{YYYY}_{MM}.xlsx', 'min_linhas': 4, 'pastas': []
        },
        {
            'nome': 'BI Painel dos Médicos', 'categoria': 'bi', 'prio': 'P1', 'regra': '',
            'log': 'log_execucao_painel_medicos_{YYYY}_{MM}.xlsx', 'min_linhas': 8,
            'hora_inicio': '07:10', 'hora_fim': '11:30', 'pastas': []
        },
        {
            'nome': 'BI Farol ICH (DGE)', 'categoria': 'bi', 'prio': 'P1', 'regra': '',
            'log': 'log_execucao_farol_de_indice_consultas_por_hora_ich_{YYYY}_{MM}.xlsx', 'min_linhas': 4,
            'hora_inicio': '14:50', 'hora_fim': '15:10',
            'pastas': [
                F('BASE CLIENTE OCULTO (DGE)', B7 + r'BI_INDICADORES_CRED\CLIENTE_OCULTO', 1, 'BASE_CLIENTE_OCULTO', True, '', conf_mensal=True, hora_limite='10:00')
            ]
        },
        {
            'nome': 'BASES Relatório de Consultas Disponibilizadas', 'categoria': 'bi', 'prio': 'P1', 'regra': '',
            'log': 'log_execucao_consultas_disponibilizadas_{YYYY}_{MM}.xlsx', 'min_linhas': 5,
            'hora_inicio': '02:00', 'hora_fim': '03:30', 'pastas': []
        },
        {
            'nome': 'BI Controle Agendas', 'categoria': 'bi', 'prio': 'P1', 'regra': '',
            'log': 'log_controle_agendas_{YYYY}_{MM}.xlsx', 'min_linhas': 5,
            'hora_inicio': '16:50', 'hora_fim': '17:30', 'pastas': []
        },
        {
            'nome': 'INDICADORES REDE CRED (Diário e Mensal)', 'categoria': 'base', 'prio': 'P1', 'regra': '',
            'log': '', 'min_linhas': 0,
            'pastas': [
                F('CLIENTE_OCULTO HISTORICO (Até 10h)', B7 + r'BI_INDICADORES_CRED\CLIENTE_OCULTO\HISTORICO', 1, 'BASE_CLIENTE_OCULTO', True, '', conf_mensal=True, hora_limite='10:00'),
                F('CLIENTE_OCULTO MENSAL (Até 10h)', B7 + r'BI_INDICADORES_CRED\CLIENTE_OCULTO', 1, 'BASE_CLIENTE_OCULTO', True, '', conf_mensal=True, hora_limite='10:00')
            ]
        },

        # 🟡 P2 — Média (Padrão D-1 e Bases Médicas)
        {
            'nome': 'T9033 - Atraso Médico', 'categoria': 'base', 'prio': 'P2', 'regra': '',
            'log': '', 'min_linhas': 0, 'd_menos_1': True,
            'pastas': [
                F(
                    rotulo='ATRASO MEDICO',
                    caminho=B7 + 'BASE_PAINEL DOS MEDICOS\\BASE_ATRASO MEDICO' + MM,
                    caminho_fallback=B7 + 'BASE_PAINEL DOS MEDICOS BASE_ATRASO MEDICO' + MM,
                    esperado=1,
                    contem='T9033_',
                    so_hoje=True,
                    conf_mensal=False
                )
            ]
        },
        {
            'nome': 'T9033 - Horas Médicos', 'categoria': 'base', 'prio': 'P2', 'regra': '',
            'log': '', 'min_linhas': 0, 'd_menos_1': True,
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
            'nome': 'BI Relatório de Consultas em Transferência', 'categoria': 'bi', 'prio': 'P2', 'regra': '',
            'log': 'log_execucao_consultas_em_transferencia_{YYYY}_{MM}.xlsx', 'min_linhas': 4,
            'hora_inicio': '22:00', 'hora_fim': '22:30', 'd_menos_1': True, 'pastas': []
        },
        {
            'nome': 'BI Relatório Atraso dos Médicos', 'categoria': 'bi', 'prio': 'P2', 'regra': '',
            'log': 'log_Atraso_dos_Medicos_{YYYY}_{MM}.xlsx', 'min_linhas': 6,
            'hora_inicio': '07:30', 'hora_fim': '09:00', 'd_menos_1': True, 'pastas': []
        },
        {
            'nome': 'BI Relatório de Acompanhamento de Mensageria', 'categoria': 'bi', 'prio': 'P2', 'regra': '',
            'log': 'log_relatorio_mensageria_{YYYY}_{MM}.xlsx', 'min_linhas': 4,
            'hora_inicio': '22:00', 'hora_fim': '22:30', 'd_menos_1': True, 'pastas': []
        },
        {
            'nome': 'CAPTADOS FORA DE ALERTA', 'categoria': 'bi', 'prio': 'P2', 'regra': '',
            'log': 'log_Script_Tabela_Captados_Fora_Alerta_{YYYY}_{MM}.xlsx', 'min_linhas': 8,
            'hora_inicio': '01:00', 'hora_fim': '01:30', 'd_menos_1': True, 'pastas': []
        },
        {
            'nome': 'Base DGE (Produtividade / Tipo Pagamento)', 'categoria': 'base', 'prio': 'P2', 'regra': '',
            'log': '', 'min_linhas': 0, 'd_menos_1': True,
            'pastas': [
                F('48 - PRODUTIVIDADE V2 DGE', r'1 - RELATORIOS\48 - RELATORIO DE PRODUTIVIDADE V2\DGE', 1, '', False, '', so_mes=True),
                F('BASES DGE PRODUTIVIDADE', B7 + r'BASES DGE\PRODUTIVIDADE', 1, '', False, '', so_mes=True),
                F('BASES DGE TIPO PAGAMENTO', B7 + r'BASES DGE\TIPO PAGAMENTO', 1, '', False, '', so_mes=True)
            ]
        },
        {
            'nome': 'Base MONITORAMENTO ANS', 'categoria': 'bi', 'prio': 'P2', 'regra': '',
            'log': 'log_execucao_consultas_em_transferencia_{YYYY}_{MM}.xlsx', 'min_linhas': 4,
            'hora_inicio': '22:00', 'hora_fim': '22:30', 'd_menos_1': True, 'pastas': []
        }
    ]
}


# ==============================================================================
# 3. AUTOCONFIGURAÇÃO E RESOLUÇÃO DE CAMINHOS
# ==============================================================================
def tok(s, d=None):
    """Substitui os marcadores {YYYY}, {MM}, {DD} e {MES} pela data informada."""
    d = d or date.today()
    return (
        (s or '')
        .replace('{YYYY}', str(d.year))
        .replace('{MM}', f'{d.month:02d}')
        .replace('{DD}', f'{d.day:02d}')
        .replace('{MES}', MES[d.month - 1])
    )


def sincronizar_config_automatico():
    """Garante que o ficheiro config_monitores.json tem sempre as novas regras, preservando os contactos do WhatsApp."""
    cfg_final = json.loads(json.dumps(DEFAULT))
    if CFG.exists():
        try:
            antigo = json.loads(CFG.read_text(encoding='utf-8'))
            if 'whatsapp' in antigo and 'destinos' in antigo['whatsapp']:
                cfg_final['whatsapp']['destinos'] = antigo['whatsapp']['destinos']
        except Exception:
            pass
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
# 4B. MONITORAMENTO DOS BIs (PBIX OFICIAL - PASTA 3) + AUDITORIA PASTA 52
# ==============================================================================
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
    """Avalia se o ficheiro .pbix cumpre a regra de atualização definida no documento."""
    hoje = agora.date()
    regra = b.get('regra', 'diaria')

    if regra == 'mensal':
        ok = (mt.year, mt.month) == (hoje.year, hoje.month)
        return (
            'ok' if ok else 'err',
            f'Atualizado em {mt:%d/%m/%Y %H:%M}' if ok else f'ATRASADO: sem atualização neste mês (última: {mt:%d/%m/%Y})',
            'Atualização mensal',
            0 if ok else int((agora - mt).total_seconds() // 60)
        )

    if regra == 'ultima':
        dias = (agora - mt).days
        lim = int(b.get('limite_dias', 3))
        return (
            'ok' if dias <= lim else 'warn',
            f'Última atualização: {mt:%d/%m/%Y %H:%M} ({dias}d atrás)',
            f'Sempre verificar última atualização (alerta > {lim}d)',
            max(0, (dias - lim) * 1440)
        )

    if hoje.weekday() >= 5 and not b.get('checar_fim_de_semana', False):
        return 'ok', f'Fim de semana (última: {mt:%d/%m %H:%M})', 'Dias úteis', 0

    if regra == 'intervalo_horas':
        interv = float(b.get('intervalo_h', 2))
        lim_min = int(interv * 60 + 15)  # 2h + 15min de tolerância
        diff_min = int((agora - mt).total_seconds() // 60)
        esp = f'Atualizado a cada {int(interv)} horas (verificar sempre a última)'
        if agora.hour < 7:
            return 'ok', f'Fora da janela matinal (última: {mt:%d/%m %H:%M})', esp, 0
        if mt.date() == hoje and diff_min <= lim_min:
            return 'ok', f'Atualizado há {diff_min} min ({mt:%H:%M:%S})', esp, 0
        atraso = max(1, diff_min - int(interv * 60))
        return 'err', f'ATRASADO: última atualização há {diff_min} min ({mt:%d/%m %H:%M:%S}) — limite {int(interv)}h', esp, atraso

    if regra == 'horarios':
        tol = int(b.get('tolerancia_min', 60))
        sl = [_hora_dt(h, hoje) for h in b['horarios']]
        passadas = [s for s in sl if s + timedelta(minutes=tol) <= agora]
        prox = next((s for s in sl if s > agora), None)
        esp = 'Janelas: ' + ', '.join(b['horarios']) + (f' | próx: {prox:%H:%M}' if prox else ' | encerrado hoje')
        if not passadas:
            return 'ok', f'Aguardando 1º horário do dia ({b["horarios"][0]})', esp, 0
        ult = passadas[-1]
        if mt >= ult - timedelta(minutes=10):
            return 'ok', f'Atualizado na janela das {ult:%H:%M} (às {mt:%H:%M:%S})', esp, 0
        atraso = int((agora - ult).total_seconds() // 60)
        return 'err', f'ATRASADO: não atualizou na janela das {ult:%H:%M}. Última: {mt:%d/%m %H:%M:%S}', esp, atraso

    # Regra 'diaria' ou 'ate_hora'
    hora_lim = b.get('hora_limite', '11:00')
    tol = int(b.get('tolerancia_min', 0))
    prazo = _hora_dt(hora_lim, hoje) + timedelta(minutes=tol)
    esp = f'Atualizado diariamente até {hora_lim}' + (f' (+{tol}m)' if tol else '')
    if mt.date() == hoje:
        return 'ok', f'Atualizado hoje às {mt:%H:%M:%S}', esp, 0
    if agora < prazo:
        return 'ok', f'Aguardando atualização até {prazo:%H:%M} (última: {mt:%d/%m %H:%M})', esp, 0
    atraso = int((agora - prazo).total_seconds() // 60)
    return 'err', f'ATRASADO: limite era {prazo:%H:%M}. Última atualização: {mt:%d/%m/%Y %H:%M:%S}', esp, atraso


def avaliar_pbix(cfg):
    agora = datetime.now()
    pasta = resolver_pasta_automatica(cfg, cfg.get('pasta_pbix', ''))
    try:
        arquivos = [
            (e, datetime.fromtimestamp(e.stat().st_mtime))
            for e in pasta.iterdir()
            if e.is_file() and e.suffix.lower() in ('.pbix', '.pbit')
        ] if pasta.is_dir() else None
    except Exception:
        arquivos = None

    if arquivos is None:
        r = _item_base('PBIX · Pasta 3 (PBIX OFICIAL)', 'P0', 'pbix')
        r['status'] = 'err'
        r['erros'].append(f'Pasta dos BIs não encontrada ou inacessível: {pasta}')
        r['pbix'] = dict(arquivo=None, ultima=None, ultima_iso=None, idade_h=None, esperado='Pasta acessível', atraso_min=0, regra='diaria')
        return [r]

    itens = []
    for b in cfg.get('bis_pbix', []):
        if b.get('regra') == 'ignorar':
            continue
        b_cfg = {'hora_limite': cfg.get('hora_limite_manha', '11:00'), 'checar_fim_de_semana': cfg.get('bi_checar_fim_de_semana', False), **b}
        r = _item_base('PBIX · ' + b_cfg['nome'], b_cfg['prio'], 'pbix')
        ach = [
            (e, mt) for e, mt in arquivos
            if all(k in _norm(e.name) for k in b_cfg['busca'])
            and not any(k in _norm(e.name) for k in b_cfg.get('excluir', []))
        ]
        if not ach:
            r['status'] = 'warn'
            r['avisos'].append(f"Ficheiro .pbix não encontrado na Pasta 3 (termos: {', '.join(b_cfg['busca'])}).")
            r['pbix'] = dict(arquivo=None, ultima=None, ultima_iso=None, idade_h=None, esperado='-', atraso_min=0, regra=b_cfg['regra'])
        else:
            e, mt = max(ach, key=lambda x: x[1])
            st, motivo, esp, atraso = _regra_pbix(b_cfg, mt, agora)
            r['status'] = st
            {'err': r['erros'], 'warn': r['avisos']}.get(st, r['infos']).append(motivo)
            if len(ach) > 1:
                r['infos'].append(f'{len(ach)} ficheiros correspondem; a usar o mais recente ({e.name})')
            r['pbix'] = dict(
                arquivo=e.name,
                ultima=mt.strftime('%d/%m/%Y %H:%M:%S'),
                ultima_iso=mt.isoformat(timespec='seconds'),
                idade_h=round((agora - mt).total_seconds() / 3600, 1),
                esperado=esp,
                atraso_min=atraso,
                regra=b_cfg['regra']
            )
        itens.append(r)
    return itens


def _dia_util_anterior(d):
    d = d - timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


def auditar_logs_52(cfg, forcar=False):
    """
    Regra Pasta 52:
    - Logs do mês atual devem estar atualizados diariamente;
    - Logs de outros meses têm que estar atualizados com a última data do mês correspondente;
    - Considera log desatualizado se a última data do log for há 3 meses atrás da referência do mês.
    """
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
        if mes_hoje - (ref[0] * 12 + ref[1]) > 6 or any(i.strip() in _norm(arq.name) for i in ignorar):
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
            ref_meses = ref[0] * 12 + ref[1]
            ult_meses = (ult.year * 12 + ult.month) if ult else 0

            if ref == (hoje.year, hoje.month):
                if ult == hoje or (hoje.weekday() >= 5 and ult and ult >= util_ant):
                    st, msg = 'ok', 'Atualizado hoje (diário OK)'
                elif ult and (ref_meses - ult_meses) >= 3:
                    st, msg = 'err', f'DESATUALIZADO CRÍTICO: última data {ult_s} está 3+ meses atrás da referência'
                elif ult and ult >= util_ant:
                    st, msg = 'warn', f'Última data {ult_s} (D-1); aguarda execução de hoje'
                else:
                    st, msg = 'err', f'DESATUALIZADO: última data {ult_s}, esperado atualização diária'
            else:
                fim_mes = date(ref[0], ref[1], calendar.monthrange(*ref)[1])
                atraso_ref = (ref_meses - ult_meses) if ult else 99
                atraso_atual = (mes_hoje - ult_meses) if ult else 99
                if atraso_ref >= 3 or (ultimo_da_familia and atraso_atual >= 3):
                    st, msg = 'err', f'LOG DESATUALIZADO (3+ meses atrás): última data {ult_s} p/ ref {ref[1]:02d}/{ref[0]}'
                elif ultimo_da_familia:
                    st, msg = 'warn', f'Sem log do mês atual (último ficheiro: {ref[1]:02d}/{ref[0]}, data {ult_s})'
                elif ult and ult >= fim_mes - timedelta(days=int(cfg.get('log_tolerancia_fim_mes_dias', 2))):
                    st, msg = 'ok', f'Fechou com a última data do mês ({ult_s})'
                else:
                    st, msg = 'warn', f'Incompleto no mês {ref[1]:02d}/{ref[0]}: última data {ult_s} (esperado {fim_mes:%d/%m/%Y})'
            det.append(dict(arquivo=a['arquivo'], ref=f'{ref[1]:02d}/{ref[0]}', ultima=ult_s, status=st, motivo=msg))

    ordem = {'err': 0, 'warn': 1, 'ok': 2}
    det.sort(key=lambda x: (ordem[x['status']], x['arquivo']))
    r['logs_detalhe'] = det[:100]
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
        r['avisos'].append(f'{len(warns)} log(s) com alerta na Pasta 52')
    r['infos'].append(f'{len(det)} ficheiros de log auditados na Pasta 52')
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
            script = next((t for t in txt if re.search(r'\.(py|yxmd|csv|xlsx|parquet)$', t, re.I)), txt[2] if len(txt) > 2 else txt[0])
            saida.append(dict(
                d=d, ini=ini, fim=fim, ini_hms=ini_hms, fim_hms=fim_hms,
                dur=dur, st=txt[si].upper()[:3], script=script, msg=' | '.join(txt[si + 1:])
            ))
    wb.close()
    return saida


def avaliar(m, cfg):
    agora = datetime.now()
    hoje = agora.date()
    ontem_util = _dia_util_anterior(hoje)
    aceita_d1 = m.get('d_menos_1', False) or m.get('prio') == 'P2'
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

    if aceita_d1:
        r['infos'].append('Padrão P2: aceita referência D-1 conforme janela de execução')

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
                if not td and aceita_d1:
                    td = [x for x in linhas if x['d'] in (hoje - timedelta(days=1), ontem_util)]
                    if td:
                        r['infos'].append(f"Exibindo execução D-1 ({td[-1]['d']:%d/%m/%Y})")

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
                        r['avisos'].append('Nenhuma execução registada hoje (ou D-1) no log')
                else:
                    rod = [x for x in ultimo.values() if x['st'] in ('EXE', 'ROD', 'EM ')]
                    if rod:
                        r['status'] = 'run'
                        r['infos'].append('A executar: ' + ', '.join(x['script'] for x in rod))
                    for x in ultimo.values():
                        if x['st'] == 'ERR':
                            r['status'] = 'err'
                            r['erros'].append(f"{x['script']} ({x['ini_hms']}): {x['msg'] or 'erro sem mensagem'}")
                    reex = [x for x in td if x['st'] == 'ERR' and ultimo[x['script']] is not x]
                    if reex:
                        r['avisos'].append(f'{len(reex)} erro(s) já reexecutado(s) com sucesso')
                    if r['status'] == 'ok' and len(td) < m.get('min_linhas', 0) and regra not in ('aprr', 'tele') and agora.strftime('%H:%M') >= m.get('hora_fim', '00:00'):
                        r['status'] = 'warn'
                        r['avisos'].append(f"Log com {len(td)} linhas; esperado {m['min_linhas']}+")
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
        eh_aprr = (f.get('regra') == 'aprr')
        esperado_dia = aprr_info['esperado'] if eh_aprr else f['esperado']
        if f.get('dias_mes'):
            esperado_dia = f['esperado'] if hoje.day in f['dias_mes'] else 0
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
                chave_d1 = tok(f.get('contem', ''), ontem_util).lower() if aceita_d1 else ''
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

                    mod_hoje = (mtime_dt.date() == hoje) or (aceita_d1 and mtime_dt.date() in (hoje - timedelta(days=1), ontem_util))
                    tem_chave = (chave_hoje in nome_low) or (bool(chave_d1) and chave_d1 in nome_low) if chave_hoje else False

                    if chave_hoje:
                        if tem_chave or mod_hoje:
                            n_hoje += 1
                            mtimes_hoje.append(mtime_dt)
                    elif not f.get('so_hoje', True) or mod_hoje:
                        n_hoje += 1
                        mtimes_hoje.append(mtime_dt)

                if f.get('so_mes') or f.get('dias_mes'):
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
                        f"{f['rotulo']}: Falta ficheiro em {len(dias_faltantes)} dia(s) útil(eis) do mês ({', '.join(dias_faltantes[:5])})"
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

                if f.get('dias_mes') and f.get('prefixos') and (hoje.day >= min(f['dias_mes'])):
                    tok_mes = tok('{MM}{YYYY}', hoje)
                    nomes_mes = [e.name.lower() for e in pasta.iterdir() if e.is_file() and tok_mes in e.name]
                    faltam = [p for p in f['prefixos'] if not any(p.lower() in nm for nm in nomes_mes)]
                    if faltam:
                        r['status'] = 'err'
                        r['erros'].append(f"NÃO BAIXOU: {f['rotulo']} sem ficheiro de {', '.join(faltam)} (dias {f['dias_mes']})")
                if n_hoje < esperado_dia:
                    r['status'] = 'err'
                    r['erros'].append(f"NÃO BAIXOU: {f['rotulo']} tem {n_hoje} de {esperado_dia} ficheiros hoje")
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
        "Tempo Médio Processo (min)", "Tempo Médio / Ficheiro (min)", "Meta Diária / Mensal", "Alertas / Erros"
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
        if r.get('pbix'):
            metas_txt = f"PBIX: {r['pbix'].get('ultima') or 'N/D'} ({r['pbix'].get('esperado')})"
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
    cab2 = ["Hora do Registo", "Horário Exato (HH:MM:SS)", "Processo", "Categoria", "Status na Hora", "Qtd Erros", "Mensagens de Falha"]
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
    log.info('Folha de cálculo guardada em %s', arq_local_datado)

    if salvar_em_pasta_especifica:
        try:
            p_rede = Path(salvar_em_pasta_especifica)
            p_rede.mkdir(parents=True, exist_ok=True)
            wb.save(p_rede / nome_arq)
        except Exception as e:
            log.warning('Não foi possível guardar cópia na rede %s: %s', salvar_em_pasta_especifica, e)

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
        log.info('A iniciar upload do ficheiro %s para o painel Vercel (%s)...', caminho_xlsx.name, url_vercel)
        encerrar_chrome_porta(porta_vercel)

        w = cfg['whatsapp']
        cands = [
            w.get('chrome_exe', ''),
            r'C:\Program Files\Google\Chrome\Application\chrome.exe',
            r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
            os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe')
        ]
        exe = next((c for c in cands if c and os.path.exists(c)), None)
        if not exe:
            log.error('chrome.exe não encontrado para enviar folha de cálculo para a Vercel.')
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

        log.info('Ficheiro %s incluído com sucesso na aba Monitoramento da Vercel!', caminho_xlsx.name)
        return True

    except Exception as e:
        log.error('Falha ao incluir ficheiro na Vercel: %s', str(e)[:250])
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
# 7. TEMPLATE DO DASHBOARD LOCAL (HTML COM ABA DE BIs, GRÁFICOS E MODAL POP-UP)
# ==============================================================================
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
    button.danger { background: var(--red); color: #fff; border-color: var(--red); }
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
    dialog { border: 0; border-radius: 12px; width: min(860px, 95vw); padding: 0; box-shadow: 0 20px 50px rgba(0,0,0,0.35); }
    dialog::backdrop { background: rgba(18,42,68,0.72); backdrop-filter: blur(2px); }
    .dh { padding: 14px 18px; background: var(--blue-inst); color: #fff; display: flex; justify-content: space-between; align-items: center; gap: 10px; }
    .db { padding: 16px 18px; background: #fff; max-height: 78vh; overflow-y: auto; }
    .hidden { display: none !important; }
    .alert-badge { position: fixed; right: 18px; bottom: 18px; z-index: 50; background: var(--red); color: #fff; border: 0; border-radius: 99px; padding: 11px 20px; font: 700 13.5px 'IBM Plex Sans', sans-serif; cursor: pointer; box-shadow: 0 4px 16px rgba(242,29,47,.55); animation: pulso 1.5s infinite; }
    @keyframes pulso { 0%,100% { transform: scale(1); } 50% { transform: scale(1.07); } }
    #dAlert { border-top: 6px solid var(--red); width: min(760px, 95vw); }
    #dAlert .dh { background: linear-gradient(90deg, var(--red), #b91c1c); }
    #dAlert .al { border-left: 5px solid var(--c); background: var(--slate-100); border-radius: 6px; padding: 10px 12px; margin-bottom: 8px; font-size: 13px; }
    #dAlert .al b { color: var(--blue-dark); font-size: 13.5px; }
    #dAlert .al small { display: block; color: var(--red); font-weight: 600; margin-top: 3px; }
    .lt { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 11.5px; }
    .lt td, .lt th { border-bottom: 1px solid var(--slate-200); padding: 4px 6px; text-align: left; }
    .lt .s-err { color: var(--red); font-weight: 700; } .lt .s-warn { color: var(--amber-800); font-weight: 600; } .lt .s-ok { color: var(--blue-inst); }
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
        </div>
      </div>
    </div>
    <div class="bar">
      <button class="g" id="bSound" title="Ativar ou silenciar o alerta sonoro do Pop-up">🔊 Som: Ativo</button>
      <button class="p" id="bRefresh">🔄 Atualizar Agora</button>
      <button class="g" id="bVercel">☁️ Enviar p/ Vercel Agora</button>
      <button class="g" id="bRel">📊 Salvar Relatório (.xlsx)</button>
      <button class="g" id="bWa">💬 Resumo no WhatsApp</button>
    </div>
  </div>
  <nav class="nav-tabs" id="tabs">
    <button class="nav-tab active" data-tab="all">Visão Geral (Todos)</button>
    <button class="nav-tab" data-tab="pbix">📊 Monitoramento de BIs (PBIX + Pasta 52)</button>
    <button class="nav-tab" data-tab="bi">Fluxos de Execução BI</button>
    <button class="nav-tab" data-tab="base">Bases de Dados (Diário + Mensal)</button>
    <button class="nav-tab" data-tab="charts">Gráficos & Evolução (Hora / Dia)</button>
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
      <strong>Manhã:</strong> até 11h | <strong>APRR/Vagas:</strong> 2h em 2h | <strong>Cliente Oculto:</strong> 10h | <strong>P2:</strong> D-1
    </div>
  </div>

  <div class="sum" id="sum"></div>

  <div class="charts" id="chartsBox">
    <div class="box"><h3>Erros e Sucessos por Hora/Minuto (Hoje)</h3><canvas id="cHora" height="170"></canvas></div>
    <div class="box"><h3>Evolução Diária (Sucessos vs Erros por Dia)</h3><canvas id="cDia" height="170"></canvas></div>
    <div class="box"><h3>Tempo Médio de Execução por Processo (min)</h3><canvas id="c1" height="170"></canvas></div>
    <div class="box"><h3>Tempo Médio por Arquivo Baixado (min / arquivo)</h3><canvas id="cArq" height="170"></canvas></div>
  </div>

  <div class="charts hidden" id="pbixBox">
    <div class="box"><h3>Proporção de Status dos BIs (Tempo Real)</h3><canvas id="cPbDonut" height="180"></canvas></div>
    <div class="box"><h3>Situação dos BIs por Prioridade (P0, P1 e P2)</h3><canvas id="cPbPrio" height="180"></canvas></div>
    <div class="box"><h3>Horas desde a última atualização de cada BI (.pbix)</h3><canvas id="cPbIdade" height="210"></canvas></div>
    <div class="box"><h3>Atualizações dos BIs por Hora do Dia (Hoje)</h3><canvas id="cPbHora" height="210"></canvas></div>
  </div>

  <div class="grid" id="grid"></div>
</main>

<button id="alertBadge" class="alert-badge hidden"></button>

<dialog id="dAlert">
  <div class="dh">
    <b>🚨 ALERTA DE MONITORAMENTO: P0 / BIs FORA DO LIMITE</b>
    <div style="display:flex;gap:6px">
      <button id="bModalSound" style="padding:5px 10px;font-size:12px">🔊 Testar / Alternar Som</button>
      <button id="bAlertOk" style="padding:5px 10px;font-size:12px"> Silenciar 30 min</button>
    </div>
  </div>
  <div class="db">
    <p style="margin:0 0 10px;font-size:13px;color:var(--slate-600)">
      Os seguintes itens críticos (<strong>P0</strong>) ou painéis (<strong>BIs</strong>) ultrapassaram o horário limite ou apresentaram erro:
    </p>
    <div id="alertList"></div>
    <div class="bar" style="margin-top:14px;justify-content:space-between">
      <button class="p" id="bAlertBis">📊 Ver Aba Monitoramento de BIs</button>
      <button id="bAlertClose">Fechar Janela</button>
    </div>
  </div>
</dialog>

<dialog id="dWa">
  <div class="dh"><b>Resumo para WhatsApp (A cada 6 horas)</b><button id="bWc">Fechar</button></div>
  <div class="db">
    <p style="font-size:12px;color:var(--slate-600);margin-top:0">O Python envia automaticamente a cada 6 horas (1 mensagem por contacto, fechando e reiniciando o navegador).</p>
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
const L = { ok: 'ATUALIZADO / OK', err: 'ATRASADO / ERRO', warn: 'ALERTA', run: 'RODANDO', nd: 'PRÉ-CADASTRADO' };
const CAT = { bi: 'Fluxo BI (Pasta 52)', base: 'Base de Dados (Pasta 7)', pbix: 'Painel BI (Pasta 3 - PBIX)', logs: 'Auditoria Pasta 52' };

let activeTab = 'all';
let charts = {};
let secLeft = 60;
let soundEnabled = localStorage.getItem('optimusSound') !== 'off';

function updateSoundBtn() {
  const txt = soundEnabled ? '🔊 Som: Ativo' : '🔇 Som: Mudo';
  $('#bSound').textContent = txt;
  if ($('#bModalSound')) $('#bModalSound').textContent = txt;
}
updateSoundBtn();

$('#bSound').onclick = () => {
  soundEnabled = !soundEnabled;
  localStorage.setItem('optimusSound', soundEnabled ? 'on' : 'off');
  updateSoundBtn();
  if (soundEnabled) beep();
};
$('#bModalSound').onclick = () => $('#bSound').click();

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
  $('#hourSel').insertAdjacentHTML('beforeend', `<option value="${h}">Estado às ${h} (Registado às ${s.salvo_as} - ${errs} erro(s))</option>`);
});

$('#hourSel').onchange = render;
$$('#tabs .nav-tab').forEach(b => b.onclick = () => {
  $$('#tabs .nav-tab').forEach(x => x.classList.remove('active'));
  b.classList.add('active');
  activeTab = b.dataset.tab;
  render();
});

$('#bRefresh').onclick = async () => {
  $('#bRefresh').disabled = true;
  $('#bRefresh').textContent = '⏳ Atualizando...';
  $('#last').textContent = 'Atualizado em: ' + new Date().toLocaleDateString('pt-BR') + ' ' + fmtHMS();
  try { await fetch('/api/atualizar'); } catch(e) {}
  location.reload();
};

$('#bVercel').onclick = async () => {
  $('#bVercel').disabled = true;
  $('#bVercel').textContent = '⏳ Enviando p/ Vercel...';
  try {
    const r = await fetch('/api/vercel');
    const j = await r.json();
    alert('Planilha gerada e enviada para a aba Monitoramento da Vercel!\nArquivo: ' + j.arquivo);
  } catch(e) {
    alert('O upload automático ocorre a cada 30 minutos pelo Python.');
  }
  $('#bVercel').disabled = false;
  $('#bVercel').textContent = '☁️ Enviar p/ Vercel Agora';
};

$('#bRel').onclick = async () => {
  try {
    const r = await fetch('/api/relatorio');
    const j = await r.json();
    alert('Relatório .xlsx salvo em:\n' + j.arquivo);
  } catch(e) {
    alert('Relatório salvo automaticamente na pasta configurada!');
  }
};

function render() {
  const hVal = $('#hourSel').value;
  let list = D.r;
  if (hVal !== 'now' && snaps[hVal]) {
    list = snaps[hVal].itens;
    $('#snapInfo').innerHTML = `<b style="color:var(--orange)">Exibindo fotografia das ${hVal} (salva às ${snaps[hVal].salvo_as})</b>`;
  } else {
    $('#snapInfo').textContent = 'Pastas 1, 3, 7 e 52 lidas automaticamente pelo Python.';
  }

  const filtered = list.filter(r => {
    if (activeTab === 'bi') return r.categoria === 'bi';
    if (activeTab === 'pbix') return r.categoria === 'pbix' || r.categoria === 'logs';
    if (activeTab === 'base') return r.categoria === 'base';
    return true;
  });

  const c = { ok: 0, err: 0, warn: 0, run: 0 };
  filtered.forEach(r => { if (c[r.status] != null) c[r.status]++; });
  $('#sum').innerHTML = [
    ['Monitorados na Aba', filtered.length, 'var(--blue-inst)'],
    ['Atualizados (OK)', c.ok, 'var(--ok)'],
    ['Em Execução', c.run, 'var(--blue-dark)'],
    ['Em Alerta', c.warn, 'var(--gold)'],
    ['Atrasados / Erro', c.err, 'var(--red)']
  ].map(([a, b, col]) => `<div class="k" style="--c:${col}"><b>${b}</b><span>${a}</span></div>`).join('');

  $('#chartsBox').classList.toggle('hidden', activeTab === 'bi' || activeTab === 'base' || activeTab === 'pbix');
  $('#pbixBox').classList.toggle('hidden', activeTab !== 'pbix');
  $('#grid').classList.toggle('hidden', activeTab === 'charts');

  $('#grid').innerHTML = [...filtered].sort((a, b) => a.prio.localeCompare(b.prio)).map(r => {
    const w = [];
    if (r.linhas_hoje != null) {
      w.push(
        ['Execuções (log)', `${r.linhas_hoje} (${r.ok_hoje} ok / ${r.err_hoje} erro)`],
        ['Último fim (HH:MM:SS)', r.ultimo_fim || '-'],
        ['Tempo médio processo', r.media_min == null ? '-' : r.media_min.toFixed(2) + ' min']
      );
    }
    if (r.media_arq_min != null || (r.pastas && r.pastas.length)) {
      w.push(['Tempo médio / arquivo', r.media_arq_min == null ? '-' : r.media_arq_min.toFixed(2) + ' min/arq']);
    }

    r.pastas.forEach(f => {
      const tArq = f.tempo_medio_arq != null ? ` (${f.tempo_medio_arq} min/arq)` : '';
      w.push([
        f.rotulo + ' (Meta Diária)',
        f.erro ? `<b style="color:var(--red)">${f.erro}</b>` : `<b>${f.n}/${f.esperado} hoje</b>${tArq} ` + (f.n >= f.esperado ? '✔' : '✖')
      ]);
      if (f.conf_mensal) {
        w.push([
          f.rotulo + ' (Meta Mensal Útil)',
          `${f.n_mes}/${f.esperado_mes} no mês (${f.dias_uteis_ok}/${f.total_dias_uteis} dias úteis) ` + (f.dias_uteis_ok >= f.total_dias_uteis ? '✔' : '⚠️')
        ]);
      }
    });
    if (r.pbix) {
      w.push(
        ['Arquivo .pbix (Pasta 3)', r.pbix.arquivo || '<b>não localizado</b>'],
        ['Última atualização', `<b>${r.pbix.ultima || '-'}</b>`],
        ['Regra de horário', r.pbix.esperado || '-']
      );
      if (r.pbix.arquivo) {
        w.push(
          ['Tempo desde atualização', r.pbix.idade_h + ' h'],
          ['Atraso acumulado', r.pbix.atraso_min ? `<b style="color:var(--red)">${r.pbix.atraso_min} min</b>` : '<b style="color:var(--ok)">No prazo</b>']
        );
      }
    }
    w.push(['Verificado em', `<b>${r.quando}</b>`]);

    return `<div class="m ${r.status}">
      <h2><span>${r.nome}</span><span class="chip">${L[r.status]}</span></h2>
      <div class="meta">
        <span class="p-tag p-${r.prio}">Prioridade ${r.prio}</span>
        <span class="c-tag">${CAT[r.categoria] || 'Base de Dados'}</span>
      </div>
      ${w.map(x => `<div class="row"><span>${x[0]}</span><span>${x[1]}</span></div>`).join('')}
      ${(r.erros || []).map(x => `<div class="msg">${x}</div>`).join('')}
      ${(r.avisos || []).map(x => `<div class="msg w">${x}</div>`).join('')}
      ${(r.infos || []).map(x => `<div class="msg i">${x}</div>`).join('')}
      ${r.logs_detalhe ? logsTable(r.logs_detalhe) : ''}
    </div>`;
  }).join('');

  renderCharts(list);
  renderPbixCharts(list);
  checkAlerts(hVal);
}

function logsTable(ls) {
  const sh = ls.filter(l => l.status !== 'ok').slice(0, 30);
  if (!sh.length) return '<div class="msg i">Todos os logs da Pasta 52 analisados estão em dia.</div>';
  return `<details open style="margin-top:8px"><summary style="cursor:pointer;font-weight:700;color:var(--blue-inst)">Ver Logs Desatualizados / Alerta (${sh.length})</summary><table class="lt"><tr><th>Ficheiro Log</th><th>Ref.</th><th>Última Data</th><th>Diagnóstico</th></tr>` +
    sh.map(l => `<tr><td>${l.arquivo}</td><td>${l.ref}</td><td>${l.ultima}</td><td class="s-${l.status}">${l.motivo}</td></tr>`).join('') + '</table></details>';
}

function renderCharts(list) {
  const mk = (id, cfg) => { charts[id]?.destroy(); charts[id] = new Chart($('#' + id), cfg); };

  const hMap = {};
  for (let h = 5; h <= 22; h++) hMap[p2(h) + ':00'] = { ok: 0, err: 0 };
  list.forEach(r => (r.eventos_hoje || []).forEach(ev => {
    const hk = (ev.h || '00:00').slice(0, 2) + ':00';
    if (!hMap[hk]) hMap[hk] = { ok: 0, err: 0 };
    if (ev.st === 'SUC') hMap[hk].ok++;
    if (ev.st === 'ERR') hMap[hk].err++;
  }));
  const hKeys = Object.keys(hMap).sort();
  mk('cHora', {
    data: {
      labels: hKeys,
      datasets: [
        { type: 'line', label: 'Linha de Erros', data: hKeys.map(k => hMap[k].err), borderColor: '#F21D2F', backgroundColor: '#F21D2F', tension: 0.3 },
        { type: 'bar', label: 'Sucessos', data: hKeys.map(k => hMap[k].ok), backgroundColor: '#0339A6' },
        { type: 'bar', label: 'Erros', data: hKeys.map(k => hMap[k].err), backgroundColor: '#F24405' }
      ]
    },
    options: { animation: false }
  });

  const dayMap = {};
  (D.hist || []).forEach(x => { dayMap[x.dia] = x; });
  const dKeys = Object.keys(dayMap).sort().slice(-15);
  mk('cDia', {
    type: 'line',
    data: {
      labels: dKeys,
      datasets: [
        { label: 'Processos OK no Dia', data: dKeys.map(k => dayMap[k].ok), borderColor: '#0339A6', backgroundColor: '#0339A6', tension: 0.25 },
        { label: 'Processos com Erro no Dia', data: dKeys.map(k => dayMap[k].err), borderColor: '#F21D2F', backgroundColor: '#F21D2F', tension: 0.25 }
      ]
    },
    options: { animation: false }
  });

  const lg = list.filter(r => r.media_min != null);
  mk('c1', {
    type: 'bar',
    data: { labels: lg.map(r => r.nome), datasets: [{ data: lg.map(r => r.media_min), backgroundColor: '#0339A6' }] },
    options: { indexAxis: 'y', plugins: { legend: { display: false } }, animation: false }
  });

  const fa = list.filter(r => r.media_arq_min != null);
  mk('cArq', {
    type: 'bar',
    data: { labels: fa.map(r => r.nome), datasets: [{ data: fa.map(r => r.media_arq_min), backgroundColor: '#F2B705' }] },
    options: { plugins: { legend: { display: false } }, animation: false }
  });
}

function renderPbixCharts(list) {
  const mk = (id, cfg) => { charts[id]?.destroy(); charts[id] = new Chart($('#' + id), cfg); };
  const COR = { ok: '#10B981', warn: '#F2B705', err: '#F21D2F', run: '#0339A6', nd: '#94A3B8' };
  const allBis = list.filter(r => r.categoria === 'pbix' || r.categoria === 'bi');

  const nOk = allBis.filter(r => r.status === 'ok').length;
  const nWarn = allBis.filter(r => r.status === 'warn').length;
  const nErr = allBis.filter(r => r.status === 'err').length;
  mk('cPbDonut', {
    type: 'doughnut',
    data: {
      labels: ['Atualizados (OK)', 'Em Alerta', 'Atrasados / Erro'],
      datasets: [{ data: [nOk, nWarn, nErr], backgroundColor: ['#10B981', '#F2B705', '#F21D2F'] }]
    },
    options: { animation: false, plugins: { legend: { position: 'bottom' } } }
  });

  const pr = ['P0', 'P1', 'P2'];
  const cnt = (p, s) => allBis.filter(r => r.prio === p && r.status === s).length;
  mk('cPbPrio', {
    type: 'bar',
    data: {
      labels: pr,
      datasets: [
        { label: 'OK', data: pr.map(p => cnt(p, 'ok')), backgroundColor: '#10B981' },
        { label: 'Alerta', data: pr.map(p => cnt(p, 'warn')), backgroundColor: '#F2B705' },
        { label: 'Atrasado / Erro', data: pr.map(p => cnt(p, 'err')), backgroundColor: '#F21D2F' }
      ]
    },
    options: { scales: { x: { stacked: true }, y: { stacked: true } }, animation: false }
  });

  const px = list.filter(r => r.categoria === 'pbix' && r.pbix && r.pbix.arquivo);
  mk('cPbIdade', {
    type: 'bar',
    data: {
      labels: px.map(r => r.nome.replace('PBIX · ', '')),
      datasets: [{ label: 'Horas desde última atualização', data: px.map(r => r.pbix.idade_h), backgroundColor: px.map(r => COR[r.status]) }]
    },
    options: { indexAxis: 'y', plugins: { legend: { display: false } }, animation: false }
  });

  const hh = Array(24).fill(0);
  px.forEach(r => {
    const d = new Date(r.pbix.ultima_iso);
    if (d.toDateString() === new Date().toDateString()) hh[d.getHours()]++;
  });
  mk('cPbHora', {
    type: 'bar',
    data: {
      labels: hh.map((_, i) => p2(i) + 'h'),
      datasets: [{ label: 'BIs atualizados nesta hora (hoje)', data: hh, backgroundColor: '#0339A6' }]
    },
    options: { animation: false }
  });
}

function beep() {
  if (!soundEnabled) return;
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    [880, 660, 880].forEach((freq, idx) => {
      const o = ctx.createOscillator(), g = ctx.createGain();
      o.connect(g); g.connect(ctx.destination);
      o.type = 'sawtooth'; o.frequency.value = freq; g.gain.value = 0.08;
      o.start(ctx.currentTime + idx * 0.22);
      o.stop(ctx.currentTime + idx * 0.22 + 0.18);
    });
  } catch (e) {}
}

function getCriticalAlerts() {
  return D.r.filter(r =>
    r.status === 'err' && (r.prio === 'P0' || r.categoria === 'pbix' || r.categoria === 'bi' || r.categoria === 'logs')
  ).sort((a, b) => a.prio.localeCompare(b.prio));
}

function ackAlert() {
  const al = getCriticalAlerts();
  localStorage.setItem('optimusAck', JSON.stringify({
    sig: al.map(r => r.nome + '|' + (r.erros[0] || '')).sort().join('#'),
    ts: Date.now()
  }));
}

function checkAlerts(hVal) {
  const al = getCriticalAlerts();
  const badge = $('#alertBadge');
  badge.classList.toggle('hidden', !al.length);
  badge.textContent = `🚨 ${al.length} Alerta(s) P0 / BI`;
  if (!al.length) { localStorage.removeItem('optimusAck'); return; }

  const warns = D.r.filter(r => r.status === 'warn').length;
  $('#alertList').innerHTML = al.slice(0, 15).map(r =>
    `<div class="al" style="--c:${r.prio === 'P0' ? 'var(--red)' : r.prio === 'P1' ? 'var(--orange)' : 'var(--gold)'}">
      <b>[${r.prio}] ${r.nome}</b>
      ${(r.erros || []).slice(0, 3).map(e => `<small>❌ ${e}</small>`).join('')}
    </div>`
  ).join('') +
  (al.length > 15 ? `<p>... e mais ${al.length - 15} alerta(s).</p>` : '') +
  (warns ? `<p style="font-size:12px;color:var(--slate-600)">Há também ${warns} item(ns) em estado de atenção (amarelo).</p>` : '') +
  `<p style="font-size:11px;color:var(--slate-600);margin-bottom:0">Leitura efetuada em: ${D.gerado}</p>`;

  const sig = al.map(r => r.nome + '|' + (r.erros[0] || '')).sort().join('#');
  const ack = JSON.parse(localStorage.getItem('optimusAck') || 'null');
  const reabrir = !ack || ack.sig !== sig || (Date.now() - ack.ts) > 30 * 60000;
  if (reabrir && hVal === 'now' && !$('#dAlert').open) {
    $('#dAlert').showModal();
    beep();
  }
}

$('#alertBadge').onclick = () => { $('#dAlert').showModal(); beep(); };
$('#bAlertOk').onclick = () => { ackAlert(); $('#dAlert').close(); };
$('#bAlertClose').onclick = () => $('#dAlert').close();
$('#dAlert').addEventListener('cancel', ackAlert);
$('#bAlertBis').onclick = () => {
  ackAlert();
  $('#dAlert').close();
  document.querySelector('#tabs [data-tab="pbix"]').click();
};

$('#last').textContent = 'Último monitoramento registado: ' + D.gerado;
setInterval(() => {
  $('#liveClock').textContent = fmtHMS(new Date());
  secLeft = Math.max(0, secLeft - 1);
  $('#cd').textContent = secLeft + 's';
  if (secLeft === 0 && !$('#dAlert').open && !$('#dWa').open) location.reload();
}, 1000);

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
    ativos = [x for x in res if x['status'] != 'nd']
    c = lambda s: sum(x['status'] == s for x in ativos)
    agora_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    hist.append(dict(
        dia=date.today().isoformat(),
        h=datetime.now().strftime('%H:%M:%S'),
        ok=c('ok'),
        warn=c('warn'),
        err=c('err'),
        bi_ok=sum(1 for x in ativos if x.get('categoria') == 'pbix' and x['status'] == 'ok'),
        bi_err=sum(1 for x in ativos if x.get('categoria') == 'pbix' and x['status'] == 'err')
    ))
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
    conteudo_html = HTML.replace('__DATA__', json.dumps(dados, ensure_ascii=False))

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
# 8. WHATSAPP A CADA 6 HORAS (360 MIN) - FECHANDO E REINICIANDO POR CONTACTO
# ==============================================================================
ICON = {'ok': '✅', 'err': '❌', 'warn': '⚠️', 'run': '🔄', 'nd': '➖'}


def resumo(res, titulo='Monitoramento OptimusBI'):
    ativos = [x for x in res if x['status'] != 'nd']
    c = lambda s: sum(x['status'] == s for x in ativos)
    t = (
        f"*{titulo}* - {datetime.now():%d/%m/%Y %H:%M:%S}\n"
        f"✅ OK: {c('ok')} | ⚠️ Alerta: {c('warn')} | ❌ Erro: {c('err')}"
        + (f" | 🔄 Rodando: {c('run')}" if c('run') else '') + '\n'
    )
    for x in sorted(ativos, key=lambda x: x['prio']):
        if x.get('categoria') == 'pbix' and x['status'] == 'ok':
            continue
        t += f"\n{ICON[x['status']]} [{x['prio']}] {x['nome']}"
        for e in x['erros']:
            t += f"\n   - {e}"
        if x['status'] == 'warn':
            for a in x['avisos'][:1]:
                t += f"\n   - {a}"
    return t


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
    cands = [
        w.get('chrome_exe', ''),
        r'C:\Program Files\Google\Chrome\Application\chrome.exe',
        r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
        os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe')
    ]
    exe = next((c for c in cands if c and os.path.exists(c)), None)
    if not exe:
        log.error('chrome.exe não encontrado.')
        return None

    perfil = w.get('perfil_chrome') or r'C:\chrome_whatsapp_monitor'
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


def enviar_whatsapp(cfg, texto):
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

    for idx, d in enumerate(destinos, start=1):
        proc = None
        try:
            log.info('[%s/%s] A iniciar janela do navegador para %s...', idx, len(destinos), d['nome'])
            encerrar_chrome_porta(porta)
            proc = abrir_chrome_limpo(cfg)
            if not proc:
                ok_all = False
                continue

            op = Options()
            op.debugger_address = f"127.0.0.1:{porta}"
            drv = webdriver.Chrome(options=op)

            drv.get(f"https://web.whatsapp.com/send?phone={d['numero']}&text={quote(texto)}")
            caixa = WebDriverWait(drv, 90).until(
                EC.presence_of_element_located((By.XPATH, "//footer//div[@contenteditable='true']"))
            )
            time.sleep(2)
            caixa.send_keys(Keys.ENTER)
            time.sleep(5)
            log.info('WhatsApp enviado para %s (1 mensagem). A fechar janela...', d['nome'])
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
    ultima_tentativa = (
        datetime.fromisoformat(est['ultima_tentativa_resumo'])
        if est.get('ultima_tentativa_resumo') else None
    )
    ativos = [x for x in res if x['status'] != 'nd']
    erros = sorted(f"{x['nome']}|{e}" for x in ativos for e in x['erros'])
    sig = '\n'.join(erros)

    pode_enviar = ultimo is None or (agora - ultimo) >= timedelta(minutes=freq_min)
    pode_tentar = (
        ultima_tentativa is None
        or (agora - ultima_tentativa) >= timedelta(minutes=freq_min)
    )
    if pode_enviar and pode_tentar:
        titulo = '🚨 Monitoramento OptimusBI (6h)' if erros else 'Monitoramento OptimusBI (6h)'
        # Registra a tentativa antes de abrir o navegador, evitando novas
        # aberturas a cada ciclo de 60 segundos quando o envio falha.
        est['ultima_tentativa_resumo'] = agora.isoformat()
        ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')
        if enviar_whatsapp(cfg, resumo(res, titulo)):
            est['ultimo_resumo'] = agora.isoformat()
            est['ultima_assinatura'] = sig
            ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')


# ==============================================================================
# 9. SERVIDOR WEB EMBUTIDO E CICLO PRINCIPAL (SEM DUPLICAÇÃO)
# ==============================================================================
ULTIMO_ESTADO_GLOBAL = {'cfg': None, 'res': [], 'snaps': {}, 'hist': []}


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
        print(f"[OptimusBI] Relatório diário .xlsx guardado em: {arq_salvo}")


def ciclo(enviar, forcar_relatorio=False, forcar_vercel=False):
    cfg = sincronizar_config_automatico()
    res = []
    for m in cfg['monitores']:
        try:
            res.append(avaliar(m, cfg))
        except Exception:
            log.error('Erro a avaliar %s\n%s', m['nome'], traceback.format_exc())
            res.append(dict(
                nome=m['nome'], categoria=m.get('categoria', 'bi'), prio=m['prio'], status='err',
                erros=['Falha interna na verificação'], avisos=[], infos=[], pastas=[],
                quando=datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                linhas_hoje=None, ok_hoje=None, err_hoje=None, media_min=None, media_arq_min=None, ultimo_fim=None
            ))

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
        rotina_whatsapp(cfg, res, enviar)
    except Exception:
        log.error('Falha na rotina WhatsApp\n%s', traceback.format_exc())

    return cfg


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--uma-vez', action='store_true')
    ap.add_argument('--sem-whatsapp', action='store_true')
    ap.add_argument('--testar-whatsapp', action='store_true')
    ap.add_argument('--gerar-relatorio', action='store_true')
    ap.add_argument('--enviar-vercel', action='store_true', help='Força geração do .xlsx e upload imediato na Vercel')
    a = ap.parse_args()

    if a.testar_whatsapp:
        sucesso = enviar_whatsapp(sincronizar_config_automatico(), 'Teste Monitoramento OptimusBI ✅ (1 mensagem por contacto)')
        print('Enviado!' if sucesso else 'Falhou - veja monitor.log')
        sys.exit()

    cfg_ini = sincronizar_config_automatico()
    if not a.uma_vez and not a.gerar_relatorio and not a.enviar_vercel:
        porta_web = int(cfg_ini.get('porta_servidor_web', 8080))
        if iniciar_servidor_web(porta_web) and cfg_ini.get('abrir_navegador_ao_iniciar', True):
            webbrowser.open(f'http://localhost:{porta_web}')

    while True:
        ini = time.time()
        try:
            cfg = ciclo(not a.sem_whatsapp, forcar_relatorio=a.gerar_relatorio, forcar_vercel=a.enviar_vercel)
        except Exception:
            log.error('Erro no ciclo\n%s', traceback.format_exc())
            cfg = DEFAULT
        if a.uma_vez or a.gerar_relatorio or a.enviar_vercel:
            break
        tempo_espera = max(10, int(cfg.get('intervalo_seg', 60)) - (time.time() - ini))
        time.sleep(tempo_espera)