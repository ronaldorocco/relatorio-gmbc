"""
Bot Interativo GMBC — Responde consultas no Telegram
Roda via GitHub Actions a cada 5 minutos.

Comandos disponíveis:
  centro              → consulta bairro (texto livre)
  /bairro centro      → consulta bairro
  /tipo furto         → consulta por tipo de crime
  /turno noite        → consulta por turno
  /resumo             → resumo geral
  /bairros            → lista todos os bairros
  /ajuda              → lista de comandos
"""

import sys, json, math, warnings
from datetime import datetime, timezone, timedelta
from collections import Counter
import urllib.request, urllib.error
import os as _os

warnings.filterwarnings('ignore')

BRT = timezone(timedelta(hours=-3))

BOT_TOKEN       = _os.environ.get('BOT_TOKEN',       '8971067969:AAF73XtvvHyhkb_KX0dC3Tny6DQ6DtRdjjM').strip()
CHAT_ID         = _os.environ.get('CHAT_ID',         '1931364088').strip()
GOOGLE_DRIVE_ID = _os.environ.get('GOOGLE_DRIVE_ID', '1w_4WgORfWrxonI-tL6uKOkoCZJQ9K5VN').strip()

# ── Mapas de normalização ────────────────────────────────────────────────────
DIA_MAP = {
    'SEGUNDA':'Segunda','TERÇA':'Terça','TERCA':'Terça','QUARTA':'Quarta',
    'QUINTA':'Quinta','SEXTA':'Sexta','SABADO':'Sábado','SÁBADO':'Sábado','DOMINGO':'Domingo',
}
BAIRRO_MAP = {
    'BARRA SUL':'Barra Sul','SÃO J. TADEU':'São J. Tadeu','SAO J. TADEU':'São J. Tadeu',
    'N. ESPERANÇA':'N. Esperança','N. ESPERANCA':'N. Esperança','MUNICIPIOS':'Municípios',
    'NAÇÕES':'Nações','NACOES':'Nações','PONTAL NORTE':'Pontal Norte',
    'VILA REAL':'Vila Real','PIONEIROS':'Pioneiros','ARIRIBA':'Ariribá',
    'CENTRO':'Centro','ESTADOS':'Estados','BARRA':'Barra',
}
MES_PARA_NUM = {
    'JANEIRO':'01','FEVEREIRO':'02','MARCO':'03','MARÇO':'03','ABRIL':'04','MAIO':'05',
    'JUNHO':'06','JULHO':'07','AGOSTO':'08','SETEMBRO':'09','OUTUBRO':'10',
    'NOVEMBRO':'11','DEZEMBRO':'12',
}
TIPO_MAP = {'FURTO':'Furto','ROUBO':'Roubo'}
ORDEM_TURNO = ['Madrugada','Manhã','Tarde','Noite']


def norm(v, mapa):
    if not v or str(v).strip().upper() in ('','NAN'): return ''
    u = str(v).strip().upper()
    return mapa.get(u, str(v).strip().title())

def norm_tipo(v):
    if not v or str(v).strip().upper() in ('','NAN'): return ''
    v = str(v).strip().upper()
    if 'TENTATIVA' in v and 'ROUBO' in v: return 'Tentativa de Roubo'
    if 'TENTATIVA' in v and 'FURTO' in v: return 'Tentativa de Furto'
    if 'ARROMBAMENTO' in v: return 'Arrombamento'
    return TIPO_MAP.get(v, v.title())

def calcular_turno(hora_val):
    try:
        h = hora_val.hour if hasattr(hora_val,'hour') else int(str(hora_val).strip().split(':')[0])
        if  6<=h<=11: return 'Manhã'
        if 12<=h<=17: return 'Tarde'
        if 18<=h<=23: return 'Noite'
        return 'Madrugada'
    except: return ''


# ── Carregar dados ───────────────────────────────────────────────────────────
def carregar_dados():
    import tempfile, os
    try: import pandas as pd
    except ImportError:
        print("ERRO: pandas nao instalado."); sys.exit(1)

    url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_DRIVE_ID}/export?format=xlsx"
    req = urllib.request.Request(url, headers={'User-Agent':'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        tmp.write(data); tmp.close()
        df = pd.read_excel(tmp.name, sheet_name='DADOS', engine='openpyxl')
        os.unlink(tmp.name)
    except Exception as e:
        print(f"ERRO ao baixar planilha: {e}"); sys.exit(1)

    df['TIPIFICACAO'] = df['TIPIFICACAO'].apply(norm_tipo)
    df['DIA_SEMANA']  = df['DIA_SEMANA'].apply(lambda v: norm(v, DIA_MAP))
    df['TURNO']       = df.apply(lambda r: calcular_turno(r['HORA']), axis=1)
    df['BAIRRO']      = df['BAIRRO'].apply(lambda v: norm(v, BAIRRO_MAP))
    df['MES_UPPER']   = df['MES'].apply(lambda v: str(v).strip().upper() if v else '')

    df = df[
        df['B.O.'].notna() &
        (df['B.O.'].astype(str).str.strip().str.upper() != 'NAN') &
        (df['B.O.'].astype(str).str.strip() != '') &
        (df['TIPIFICACAO'] != '')
    ].copy()

    def build_data(row):
        mes_num = MES_PARA_NUM.get(row['MES_UPPER'],'')
        if not mes_num: return ''
        try: return f"{int(row['ANO'])}-{mes_num}-{str(int(row['DATA'])).zfill(2)}"
        except: return ''

    df['DATA_STR'] = df.apply(build_data, axis=1)
    df['HORA_STR'] = df['HORA'].apply(lambda x: x.strftime('%H:%M') if hasattr(x,'strftime') else str(x)[:5] if str(x).strip() not in ('','nan') else '')
    df['ENDERECO'] = df['ENDEREÇO'].fillna('').astype(str)
    df['ITEM_STR'] = df['ITEM'].fillna('').astype(str).apply(lambda v: '' if v.strip().lower() in ('','nan') else v.strip().title()) if 'ITEM' in df.columns else ''
    df['MARCA_STR'] = df['MARCA_MODELO'].fillna('').astype(str).apply(lambda v: '' if v.strip().lower() in ('','nan') else v.strip()) if 'MARCA_MODELO' in df.columns else ''
    df['BO_STR'] = df['B.O.'].fillna('').astype(str)

    records = []
    for _, r in df.iterrows():
        if not r['DATA_STR']: continue
        records.append({
            'data':    r['DATA_STR'],
            'dia':     r['DIA_SEMANA'],
            'turno':   r['TURNO'],
            'tipo':    r['TIPIFICACAO'],
            'bairro':  r['BAIRRO'],
            'endereco':r['ENDERECO'],
            'hora':    r['HORA_STR'],
            'item':    r.get('ITEM_STR','') if 'ITEM_STR' in r else '',
            'marca':   r.get('MARCA_STR','') if 'MARCA_STR' in r else '',
            'bo':      r['BO_STR'],
        })
    return records


# ── Telegram API ─────────────────────────────────────────────────────────────
def get_updates(offset=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates?timeout=3"
    if offset: url += f"&offset={offset}"
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            return json.loads(resp.read().decode()).get('result', [])
    except Exception as e:
        print(f"Erro getUpdates: {e}"); return []

def send_message(chat_id, text):
    url  = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    data = json.dumps({"chat_id": chat_id, "text": text, "parse_mode": "Markdown"}).encode('utf-8')
    req  = urllib.request.Request(url, data=data, headers={'Content-Type':'application/json'})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read().decode())
            if result.get('ok'): print(f"  Resposta enviada para {chat_id}")
            else: print(f"  ERRO Telegram: {result.get('description')}")
    except Exception as e:
        print(f"  ERRO ao enviar: {e}")


# ── Consultas ─────────────────────────────────────────────────────────────────
def pct(n, total):
    return f"{round(n/total*100)}%" if total else "0%"

def consultar_bairro(records, query):
    bairros = sorted(set(r['bairro'] for r in records if r['bairro']))
    q = query.strip().lower()
    match = next((b for b in bairros if q in b.lower() or b.lower() in q), None)

    if not match:
        lista = '\n'.join(f"• {b}" for b in bairros[:15])
        return f"Bairro *'{query}'* não encontrado.\n\n*Bairros disponíveis:*\n{lista}"

    dados  = [r for r in records if r['bairro'] == match]
    total  = len(dados)
    tipos  = Counter(r['tipo']     for r in dados if r['tipo']).most_common(5)
    turnos = Counter(r['turno']    for r in dados if r['turno']).most_common()
    ruas   = Counter(r['endereco'] for r in dados if r['endereco']).most_common(5)
    horas  = Counter(r['hora'][:2] for r in dados if r['hora'] and len(r['hora'])>=2).most_common(3)
    datas  = sorted(set(r['data']  for r in dados))

    linhas = [
        f"📍 *BAIRRO: {match.upper()}*",
        f"📊 Total de ocorrências: *{total}*",
        f"📅 Período: {datas[0][8:10]}/{datas[0][5:7]} — {datas[-1][8:10]}/{datas[-1][5:7]}",
        "",
        "*🔴 Tipos de crime:*",
        *[f"  {i+1}. {t}: *{n}* ({pct(n,total)})" for i,(t,n) in enumerate(tipos)],
        "",
        "*⏰ Por turno:*",
        *[f"  • {t}: *{n}* ({pct(n,total)})" for t,n in turnos],
        "",
        "*🕐 Horários de pico:*",
        *[f"  • {h}h: {n} oc." for h,n in horas],
        "",
        "*🛣️ Ruas mais afetadas:*",
        *[f"  {i+1}. {r}: *{n}* oc." for i,(r,n) in enumerate(ruas) if r and r!='nan'],
    ]
    return '\n'.join(linhas)

def consultar_tipo(records, query):
    tipos = sorted(set(r['tipo'] for r in records if r['tipo']))
    q = query.strip().lower()
    match = next((t for t in tipos if q in t.lower() or t.lower() in q), None)

    if not match:
        lista = '\n'.join(f"• {t}" for t in tipos)
        return f"Tipo *'{query}'* não encontrado.\n\n*Tipos disponíveis:*\n{lista}"

    dados   = [r for r in records if r['tipo'] == match]
    total   = len(dados)
    bairros = Counter(r['bairro'] for r in dados if r['bairro']).most_common(5)
    turnos  = Counter(r['turno']  for r in dados if r['turno']).most_common()
    ruas    = Counter(r['endereco'] for r in dados if r['endereco']).most_common(3)
    datas   = sorted(set(r['data'] for r in dados))

    linhas = [
        f"🔴 *CRIME: {match.upper()}*",
        f"📊 Total: *{total}* ocorrências",
        f"📅 Período: {datas[0][8:10]}/{datas[0][5:7]} — {datas[-1][8:10]}/{datas[-1][5:7]}",
        "",
        "*📍 Bairros mais afetados:*",
        *[f"  {i+1}. {b}: *{n}* ({pct(n,total)})" for i,(b,n) in enumerate(bairros)],
        "",
        "*⏰ Por turno:*",
        *[f"  • {t}: *{n}* ({pct(n,total)})" for t,n in turnos],
        "",
        "*🛣️ Ruas de maior risco:*",
        *[f"  {i+1}. {r}: *{n}* oc." for i,(r,n) in enumerate(ruas) if r and r!='nan'],
    ]
    return '\n'.join(linhas)

def consultar_turno(records, query):
    turno_map = {'manha':'Manhã','manhã':'Manhã','tarde':'Tarde','noite':'Noite','madrugada':'Madrugada'}
    match = turno_map.get(query.strip().lower())
    if not match:
        return "Turnos disponíveis: *Madrugada*, *Manhã*, *Tarde*, *Noite*"

    dados   = [r for r in records if r['turno'] == match]
    total   = len(dados)
    bairros = Counter(r['bairro'] for r in dados if r['bairro']).most_common(5)
    tipos   = Counter(r['tipo']   for r in dados if r['tipo']).most_common(5)

    linhas = [
        f"⏰ *TURNO: {match.upper()}*",
        f"📊 Total: *{total}* ocorrências",
        "",
        "*📍 Bairros mais afetados:*",
        *[f"  {i+1}. {b}: *{n}* ({pct(n,total)})" for i,(b,n) in enumerate(bairros)],
        "",
        "*🔴 Crimes mais frequentes:*",
        *[f"  {i+1}. {t}: *{n}* ({pct(n,total)})" for i,(t,n) in enumerate(tipos)],
    ]
    return '\n'.join(linhas)

def consultar_resumo(records):
    total   = len(records)
    bairros = Counter(r['bairro'] for r in records if r['bairro']).most_common(3)
    tipos   = Counter(r['tipo']   for r in records if r['tipo']).most_common(3)
    turnos  = Counter(r['turno']  for r in records if r['turno']).most_common()
    datas   = sorted(set(r['data'] for r in records))
    now     = datetime.now(BRT)

    linhas = [
        f"📊 *RESUMO GERAL — GMBC*",
        f"🕐 {now.strftime('%d/%m/%Y às %H:%M')}",
        f"Total de ocorrências: *{total}*",
        f"Período: {datas[0][8:10]}/{datas[0][5:7]} — {datas[-1][8:10]}/{datas[-1][5:7]}",
        "",
        "*📍 Top bairros:*",
        *[f"  {i+1}. {b}: {n} oc." for i,(b,n) in enumerate(bairros)],
        "",
        "*🔴 Top crimes:*",
        *[f"  {i+1}. {t}: {n} oc." for i,(t,n) in enumerate(tipos)],
        "",
        "*⏰ Por turno:*",
        *[f"  • {t}: {n} ({pct(n,total)})" for t,n in turnos],
    ]
    return '\n'.join(linhas)

def busca_universal(records, query):
    """Busca em todos os campos: bairro, tipo, item, marca, turno, logradouro, dia."""
    q = query.strip().lower()
    campos = ['bairro','tipo','item','marca','turno','endereco','dia','bo']

    # Variações da busca: original, sem 's' final, sem 'es' final
    variacoes = {q}
    if q.endswith('s') and len(q) > 3:
        variacoes.add(q[:-1])     # bicicletas → bicicleta
    if q.endswith('es') and len(q) > 4:
        variacoes.add(q[:-2])     # veiculos → veiculo

    def campo_match(valor):
        v = str(valor).lower()
        return any(var in v for var in variacoes)

    matches = [r for r in records if any(campo_match(r.get(c,'')) for c in campos)]

    if not matches:
        return (
            f"Nenhuma ocorrência encontrada para *'{query}'*.\n\n"
            "Tente: nome de bairro, tipo de crime, item (bicicleta, celular...), turno ou logradouro.\n"
            "Digite /ajuda para ver todos os comandos."
        )

    total = len(matches)

    # Registros onde a busca bateu no campo ITEM (igual ao filtro do dashboard)
    matches_item = [r for r in matches if campo_match(r.get('item',''))]
    total_item = len(matches_item)

    # Descobre em quais campos houve match para exibir contexto
    campos_encontrados = [c for c in campos if any(campo_match(r.get(c,'')) for r in matches)]

    tipos   = Counter(r['tipo']    for r in matches if r['tipo']).most_common(5)
    bairros = Counter(r['bairro']  for r in matches if r['bairro']).most_common(5)
    turnos  = Counter(r['turno']   for r in matches if r['turno']).most_common()
    ruas    = Counter(r['endereco'] for r in matches if r['endereco'] and r['endereco'] != 'nan').most_common(5)
    dias    = Counter(r['dia']     for r in matches if r['dia']).most_common()
    meses   = Counter(r['data'][5:7] for r in matches if r.get('data') and len(r['data'])>=7).most_common()
    datas   = sorted(set(r['data'] for r in matches))

    MES_NOME = {'01':'Janeiro','02':'Fevereiro','03':'Março','04':'Abril','05':'Maio',
                '06':'Junho','07':'Julho','08':'Agosto','09':'Setembro','10':'Outubro',
                '11':'Novembro','12':'Dezembro'}

    # Linha de resumo
    if total_item > 0 and total_item < total:
        resumo = f"*{total_item}* como item principal  |  *{total}* total (incl. descrições)"
    else:
        resumo = f"*{total}* ocorrência(s) encontrada(s)"

    linhas = [
        f"🔍 *BUSCA: \"{query.upper()}\"*",
        f"📊 {resumo}",
        f"📅 Período: {datas[0][8:10]}/{datas[0][5:7]} — {datas[-1][8:10]}/{datas[-1][5:7]}",
    ]

    if tipos:
        linhas += ["", "*🔴 Tipificação:*"]
        linhas += [f"  {i+1}. {t}: *{n}* ({pct(n,total)})" for i,(t,n) in enumerate(tipos)]

    if bairros:
        linhas += ["", "*📍 Bairros:*"]
        linhas += [f"  {i+1}. {b}: *{n}* ({pct(n,total)})" for i,(b,n) in enumerate(bairros)]

    if turnos:
        linhas += ["", "*⏰ Turno:*"]
        linhas += [f"  • {t}: *{n}* ({pct(n,total)})" for t,n in turnos]

    if dias:
        linhas += ["", "*📆 Dia da semana:*"]
        linhas += [f"  • {d}: *{n}* ({pct(n,total)})" for d,n in dias]

    if meses:
        linhas += ["", "*🗓 Mês:*"]
        linhas += [f"  • {MES_NOME.get(m,m)}: *{n}*" for m,n in meses]

    if ruas:
        linhas += ["", "*🛣️ Ruas:*"]
        linhas += [f"  {i+1}. {r}: *{n}* oc." for i,(r,n) in enumerate(ruas)]

    return '\n'.join(linhas)


AJUDA = (
    "*🛡️ Bot GMBC — Consulta de Ocorrências*\n\n"
    "*Busca livre — digite qualquer palavra:*\n"
    "  `bicicleta` → todas as ocorrências com bicicleta\n"
    "  `celular` → todas com celular\n"
    "  `centro` → todas no bairro Centro\n"
    "  `furto` → todas por furto\n"
    "  `noite` → todas no turno noite\n"
    "  `r. 2500` → todas na Rua 2500\n\n"
    "*Comandos específicos:*\n"
    "  `/bairro centro` → estatísticas do bairro\n"
    "  `/tipo furto` → análise por tipo de crime\n"
    "  `/turno noite` → análise por turno\n"
    "  `/resumo` → resumo geral\n"
    "  `/bairros` → lista todos os bairros\n"
    "  `/tipos` → lista todos os tipos de crime\n"
    "  `/ajuda` → este menu\n\n"
    "_Resposta em até 5 minutos._"
)


# ── Processar mensagem ────────────────────────────────────────────────────────
def processar(text, chat_id, records):
    t  = text.strip()
    tl = t.lower()

    if tl in ['/start','/ajuda','/help','ajuda','help']:
        return send_message(chat_id, AJUDA)

    if tl == '/resumo':
        return send_message(chat_id, consultar_resumo(records))

    if tl == '/bairros':
        bairros = sorted(set(r['bairro'] for r in records if r['bairro']))
        return send_message(chat_id, "*📍 Bairros disponíveis:*\n" + '\n'.join(f"• {b}" for b in bairros))

    if tl == '/tipos':
        tipos = sorted(set(r['tipo'] for r in records if r['tipo']))
        return send_message(chat_id, "*🔴 Tipos de crime:*\n" + '\n'.join(f"• {t}" for t in tipos))

    if tl.startswith('/bairro '):
        return send_message(chat_id, consultar_bairro(records, t[8:]))

    if tl.startswith('/tipo '):
        return send_message(chat_id, consultar_tipo(records, t[6:]))

    if tl.startswith('/turno '):
        return send_message(chat_id, consultar_turno(records, t[7:]))

    # Busca universal em todos os campos
    return send_message(chat_id, busca_universal(records, t))


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    import time

    print("Bot GMBC iniciado — modo continuo...")
    print("Carregando dados iniciais...")
    records = carregar_dados()
    print(f"  {len(records)} registros carregados.")

    # Recarrega dados a cada 30 minutos
    ultima_carga = time.time()
    INTERVALO_RECARGA = 30 * 60  # 30 minutos

    offset = None
    print("Aguardando mensagens...\n")

    while True:
        try:
            # Recarrega dados periodicamente
            if time.time() - ultima_carga > INTERVALO_RECARGA:
                print("Recarregando dados da planilha...")
                records = carregar_dados()
                print(f"  {len(records)} registros.")
                ultima_carga = time.time()

            # Long polling — espera até 30s por novas mensagens
            url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates?timeout=30"
            if offset:
                url += f"&offset={offset}"

            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=40) as resp:
                    updates = json.loads(resp.read().decode()).get('result', [])
            except Exception as e:
                print(f"Erro polling: {e}")
                time.sleep(5)
                continue

            for upd in updates:
                offset = upd['update_id'] + 1
                msg = upd.get('message') or upd.get('edited_message')
                if not msg:
                    continue
                text = msg.get('text', '').strip()
                if not text:
                    continue
                cid  = str(msg['chat']['id'])
                nome = msg['chat'].get('first_name') or msg['chat'].get('title') or cid
                now  = datetime.now(BRT).strftime('%H:%M:%S')
                print(f"[{now}] [{nome}] '{text}'")
                processar(text, cid, records)

        except KeyboardInterrupt:
            print("\nBot encerrado.")
            break
        except Exception as e:
            print(f"Erro inesperado: {e}")
            time.sleep(5)
