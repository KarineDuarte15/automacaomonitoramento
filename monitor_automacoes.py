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
      arquivos_obrigatorios=None, prefixos=None, caminho_fallback=''):
    return dict(
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
        'porta_chrome': 9222,
        'resumo_a_cada_min': 360,  # Envia no WhatsApp apenas a cada 6 horas (360 minutos)
        'horarios_resumo': ['00:00', '06:00', '12:00', '18:00'],
        'alerta_imediato_erro': False,  # Desativado para respeitar rigorosamente a janela de 6 horas
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
                    if r['status'] == 'ok' and len(td) < m.get('min_linhas', 0) and regra not in ('aprr', 'tele'):
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
            "Painel BI" if r['categoria'] == 'bi' else "Base de Dados",
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
    Abre o Google Chrome no link https://acompanhameto-fluxo-diario-de-taref.vercel.app/,
    clica na aba Monitoramento e inclui o arquivo .xlsx gerado.
    """
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    url_vercel = cfg.get('url_painel_vercel', DEFAULT['url_painel_vercel'])
    porta_vercel = 9225  # Porta separada para não conflitar com o WhatsApp (9222)
    proc = None
    drv = None

    try:
        log.info('Iniciando upload do arquivo %s para o painel Vercel (%s)...', caminho_xlsx.name, url_vercel)
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
        encerrar_chrome_porta(porta_vercel)


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
HTML = r'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta http-equiv="refresh" content="60">
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
      <button class="p" id="bRefresh">🔄 Atualizar Agora</button>
      <button class="g" id="bVercel">☁️ Enviar p/ Vercel Agora</button>
      <button class="g" id="bRel">📊 Salvar Relatório (.xlsx)</button>
      <button class="g" id="bWa">💬 Resumo no WhatsApp</button>
    </div>
  </div>
  <nav class="nav-tabs" id="tabs">
    <button class="nav-tab active" data-tab="all">Visão Geral (Todos)</button>
    <button class="nav-tab" data-tab="bi">Painéis BI</button>
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
      <strong>Vercel (.xlsx):</strong> a cada 30 min | <strong>WhatsApp:</strong> a cada 6 horas | <strong>Relatório Final:</strong> 18:00
    </div>
  </div>

  <div class="sum" id="sum"></div>

  <div class="charts" id="chartsBox">
    <div class="box"><h3>Erros e Sucessos por Hora/Minuto (Hoje)</h3><canvas id="cHora" height="170"></canvas></div>
    <div class="box"><h3>Evolução Diária (Sucessos vs Erros por Dia)</h3><canvas id="cDia" height="170"></canvas></div>
    <div class="box"><h3>Tempo Médio de Execução por Processo (min)</h3><canvas id="c1" height="170"></canvas></div>
    <div class="box"><h3>Tempo Médio por Arquivo Baixado (min / arquivo)</h3><canvas id="cArq" height="170"></canvas></div>
  </div>

  <div class="grid" id="grid"></div>
</main>

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

let activeTab = 'all';
let charts = {};
let secLeft = 60;

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
    $('#snapInfo').textContent = 'Pastas lidas automaticamente pelo Python.';
  }

  const filtered = list.filter(r => {
    if (activeTab === 'bi') return r.categoria === 'bi';
    if (activeTab === 'base') return r.categoria === 'base';
    return true;
  });

  const c = { ok: 0, err: 0, warn: 0, run: 0 };
  filtered.forEach(r => { if (c[r.status] != null) c[r.status]++; });
  $('#sum').innerHTML = [
    ['Monitorados', filtered.length, 'var(--blue-inst)'],
    ['Status OK', c.ok, 'var(--ok)'],
    ['Rodando', c.run, 'var(--blue-dark)'],
    ['Alertas (Média)', c.warn, 'var(--gold)'],
    ['Crítico / Erro', c.err, 'var(--red)']
  ].map(([a, b, col]) => `<div class="k" style="--c:${col}"><b>${b}</b><span>${a}</span></div>`).join('');

  $('#chartsBox').classList.toggle('hidden', activeTab === 'bi' || activeTab === 'base');
  $('#grid').classList.toggle('hidden', activeTab === 'charts');

  $('#grid').innerHTML = [...filtered].sort((a, b) => a.prio.localeCompare(b.prio)).map(r => {
    const w = [];
    if (r.linhas_hoje != null) {
      w.push(
        ['Execuções hoje', `${r.linhas_hoje} (${r.ok_hoje} ok / ${r.err_hoje} erro)`],
        ['Último fim (HH:MM:SS)', r.ultimo_fim || '-'],
        ['Tempo médio processo', r.media_min == null ? '-' : r.media_min.toFixed(2) + ' min']
      );
    }
    if (r.media_arq_min != null || (r.pastas && r.pastas.length)) {
      w.push(['Tempo médio / arquivo baixado', r.media_arq_min == null ? '-' : r.media_arq_min.toFixed(2) + ' min/arq']);
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
    w.push(['Verificado em (Hora:Min:Seg)', `<b>${r.quando}</b>`]);

    return `<div class="m ${r.status}">
      <h2><span>${r.nome}</span><span class="chip">${L[r.status]}</span></h2>
      <div class="meta">
        <span class="p-tag p-${r.prio}">Prioridade ${r.prio}</span>
        <span class="c-tag">${r.categoria === 'bi' ? 'Painel BI' : 'Base de Dados'}</span>
      </div>
      ${w.map(x => `<div class="row"><span>${x[0]}</span><span>${x[1]}</span></div>`).join('')}
      ${(r.erros || []).map(x => `<div class="msg">${x}</div>`).join('')}
      ${(r.avisos || []).map(x => `<div class="msg w">${x}</div>`).join('')}
      ${(r.infos || []).map(x => `<div class="msg i">${x}</div>`).join('')}
    </div>`;
  }).join('');

  renderCharts(list);
}

function renderCharts(list) {
  const mk = (id, cfg) => { charts[id]?.destroy(); charts[id] = new Chart($('#' + id), cfg); };

  const hMap = {};
  for (let h = 5; h <= 20; h++) hMap[p2(h) + ':00'] = { ok: 0, err: 0 };
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

$('#last').textContent = 'Último monitoramento registrado: ' + D.gerado;
setInterval(() => {
  $('#liveClock').textContent = fmtHMS(new Date());
  secLeft = Math.max(0, secLeft - 1);
  $('#cd').textContent = secLeft + 's';
  if (secLeft === 0) location.reload();
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
    hist.append(dict(dia=date.today().isoformat(), h=datetime.now().strftime('%H:%M:%S'), ok=c('ok'), warn=c('warn'), err=c('err')))
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
# 8. WHATSAPP A CADA 6 HORAS (360 MIN) - FECHANDO E REINICIANDO POR CONTATO
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
            log.info('[%s/%s] Iniciando janela do navegador para %s...', idx, len(destinos), d['nome'])
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
            log.info('WhatsApp enviado para %s (1 mensagem). Fechando janela...', d['nome'])
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
        titulo = '🚨 Monitoramento OptimusBI (6h)' if erros else 'Monitoramento OptimusBI (6h)'
        if enviar_whatsapp(cfg, resumo(res, titulo)):
            est['ultimo_resumo'] = agora.isoformat()
            est['ultima_assinatura'] = sig
            ESTADO.write_text(json.dumps(est, ensure_ascii=False), encoding='utf-8')


# ==============================================================================
# 9. SERVIDOR WEB EMBUTIDO E CICLO PRINCIPAL
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
        print(f"[OptimusBI] Relatório diário .xlsx salvo em: {arq_salvo}")


def ciclo(enviar, forcar_relatorio=False, forcar_vercel=False):
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
        sucesso = enviar_whatsapp(sincronizar_config_automatico(), 'Teste Monitoramento OptimusBI ✅ (1 mensagem por contato)')
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