"""
╔══════════════════════════════════════════════════════════════════════╗
║   ENVIO AUTOMÁTICO — ANÁLISE DIÁRIA + PREVISÃO — GUARDA MUNICIPAL BC ║
║                                                                      ║
║  CONFIGURAÇÃO (faça uma vez só):                                     ║
║                                                                      ║
║  1. Crie o bot no Telegram:                                          ║
║     → Abra @BotFather no Telegram                                    ║
║     → Digite /newbot e siga as instruções                            ║
║     → Copie o TOKEN gerado e cole em BOT_TOKEN abaixo                ║
║                                                                      ║
║  2. Obtenha o CHAT_ID do grupo/supervisor:                           ║
║     → Adicione o bot ao grupo (ou inicie conversa direta)            ║
║     → Rode: python enviar_relatorio.py --get-chat-id                 ║
║     → Cole o ID em CHAT_ID abaixo                                    ║
║                                                                      ║
║  3. Agende no Windows:                                               ║
║     → Abra "Agendador de Tarefas" do Windows                         ║
║     → Criar tarefa básica → horário desejado (ex: 06h50)             ║
║     → Ação: iniciar programa → python                                ║
║     → Argumentos: "C:\...\enviar_relatorio.py"                       ║
║                                                                      ║
╚══════════════════════════════════════════════════════════════════════╝
"""

import sys
import json
import math
import warnings
from datetime import datetime
from collections import Counter
import urllib.request
import urllib.parse
import urllib.error

warnings.filterwarnings('ignore')

# ╔══════════════════════════════════════════════════════╗
# ║            CONFIGURAÇÃO — EDITE AQUI                 ║
# ╚══════════════════════════════════════════════════════╝

import os as _os

# Lê das variáveis de ambiente (GitHub Actions) ou usa os valores locais como fallback
BOT_TOKEN       = _os.environ.get('BOT_TOKEN',       "8971067969:AAF73XtvvHyhkb_KX0dC3Tny6DQ6DtRdjjM").strip()
CHAT_ID         = _os.environ.get('CHAT_ID',         "1931364088").strip()
GOOGLE_DRIVE_ID = _os.environ.get('GOOGLE_DRIVE_ID', "1w_4WgORfWrxonI-tL6uKOkoCZJQ9K5VN").strip()

ARQUIVO_EXCEL = "secretario.xlsx"   # Usado apenas se GOOGLE_DRIVE_ID estiver vazio

# ──────────────────────────────────────────────────────


DIA_MAP = {
    'SEGUNDA':'Segunda','TERÇA':'Terça','TERCA':'Terça','QUARTA':'Quarta',
    'QUINTA':'Quinta','SEXTA':'Sexta','SABADO':'Sábado','SÁBADO':'Sábado',
    'DOMINGO':'Domingo',
}
TURNO_MAP = {
    'MANHA':'Manhã','MANHÃ':'Manhã','TARDE':'Tarde',
    'NOITE':'Noite','MADRUGADA':'Madrugada',
}
MES_PARA_NUM = {
    'JANEIRO':'01','FEVEREIRO':'02','MARCO':'03','MARÇO':'03',
    'ABRIL':'04','MAIO':'05','JUNHO':'06','JULHO':'07',
    'AGOSTO':'08','SETEMBRO':'09','OUTUBRO':'10',
    'NOVEMBRO':'11','DEZEMBRO':'12',
}
BAIRRO_MAP = {
    'BARRA SUL':'Barra Sul','SÃO J. TADEU':'São J. Tadeu',
    'SAO J. TADEU':'São J. Tadeu','N. ESPERANÇA':'N. Esperança',
    'N. ESPERANCA':'N. Esperança','MUNICIPIOS':'Municípios',
    'NAÇÕES':'Nações','NACOES':'Nações','PONTAL NORTE':'Pontal Norte',
    'VILA REAL':'Vila Real','PIONEIROS':'Pioneiros','ARIRIBA':'Ariribá',
    'CENTRO':'Centro','ESTADOS':'Estados','BARRA':'Barra',
}
TIPO_MAP = {
    'FURTO':'Furto','ROUBO':'Roubo',
}
ORDEM_TURNO = ['Madrugada', 'Manhã', 'Tarde', 'Noite']


def norm(v, mapa, fallback_title=True):
    if not v or str(v).strip().upper() in ('', 'NAN'):
        return ''
    u = str(v).strip().upper()
    if u in mapa:
        return mapa[u]
    if fallback_title:
        return str(v).strip().title()
    return str(v).strip()


def norm_tipo(v):
    if not v or str(v).strip().upper() in ('', 'NAN'):
        return ''
    v = str(v).strip().upper()
    if 'TENTATIVA' in v and 'ROUBO' in v:
        return 'Tentativa de Roubo'
    if 'TENTATIVA' in v and 'FURTO' in v:
        return 'Tentativa de Furto'
    if 'ARROMBAMENTO' in v:
        return 'Arrombamento'
    return TIPO_MAP.get(v, v.title())


def calcular_turno(hora_val):
    try:
        if hasattr(hora_val, 'hour'):
            h = hora_val.hour
        elif hora_val and str(hora_val).strip():
            h = int(str(hora_val).strip().split(':')[0])
        else:
            return ''
        if  6 <= h <= 11: return 'Manhã'
        if 12 <= h <= 17: return 'Tarde'
        if 18 <= h <= 23: return 'Noite'
        return 'Madrugada'
    except Exception:
        return ''


def baixar_google_drive(file_id):
    import tempfile
    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    print(f"  Baixando planilha do Google Drive...")
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        tmp.write(data)
        tmp.close()
        print(f"  Download OK ({len(data)//1024} KB)")
        return tmp.name
    except Exception as e:
        print(f"ERRO ao baixar do Google Drive: {e}")
        print("Verifique se o arquivo esta compartilhado como 'Qualquer pessoa com o link'.")
        sys.exit(1)


def carregar_dados():
    import os

    try:
        import pandas as pd
    except ImportError:
        print("ERRO: pandas nao instalado. Execute: pip install pandas openpyxl")
        sys.exit(1)

    if GOOGLE_DRIVE_ID:
        excel_path = baixar_google_drive(GOOGLE_DRIVE_ID)
        remover_temp = True
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        excel_path = os.path.join(script_dir, ARQUIVO_EXCEL)
        remover_temp = False

    df = pd.read_excel(excel_path, sheet_name='DADOS', engine='openpyxl')

    if remover_temp:
        try:
            os.unlink(excel_path)
        except Exception:
            pass

    # Normalizar
    df['TIPIFICACAO'] = df['TIPIFICACAO'].apply(norm_tipo)
    df['DIA_SEMANA']  = df['DIA_SEMANA'].apply(lambda v: norm(v, DIA_MAP))
    df['TURNO']       = df.apply(lambda r: calcular_turno(r['HORA']), axis=1)
    df['BAIRRO']      = df['BAIRRO'].apply(lambda v: norm(v, BAIRRO_MAP))
    df['MES_UPPER']   = df['MES'].apply(lambda v: str(v).strip().upper() if v else '')

    # Filtrar linhas válidas
    df = df[
        df['B.O.'].notna() &
        (df['B.O.'].astype(str).str.strip().str.upper() != 'NAN') &
        (df['B.O.'].astype(str).str.strip() != '') &
        (df['TIPIFICACAO'] != '')
    ].copy()

    # Construir DATA_STR
    def build_data(row):
        mes_num = MES_PARA_NUM.get(row['MES_UPPER'], '')
        if not mes_num:
            return ''
        try:
            dia = str(int(row['DATA'])).zfill(2)
            ano = str(int(row['ANO']))
            return f"{ano}-{mes_num}-{dia}"
        except Exception:
            return ''

    df['DATA_STR'] = df.apply(build_data, axis=1)
    df['ENDERECO'] = df['ENDEREÇO'].fillna('').astype(str)

    records = []
    for _, r in df.iterrows():
        records.append({
            'data':    r['DATA_STR'],
            'dia':     r['DIA_SEMANA'],
            'turno':   r['TURNO'],
            'tipo':    r['TIPIFICACAO'],
            'bairro':  r['BAIRRO'],
            'endereco': r['ENDERECO'],
        })

    return [rec for rec in records if rec['data']]


def top_n(records, campo, n=5):
    c = Counter(r[campo] for r in records if r[campo])
    return c.most_common(n)


def pct(n, total):
    if not total:
        return '0%'
    return f"{round(n/total*100)}%"


def turno_atual():
    h = datetime.now().hour
    if  6 <= h <= 11: return 'Manhã'
    if 12 <= h <= 17: return 'Tarde'
    if 18 <= h <= 23: return 'Noite'
    return 'Madrugada'


def proxima_turno(t):
    prox = {'Madrugada':'Manhã','Manhã':'Tarde','Tarde':'Noite','Noite':'Madrugada'}
    return prox.get(t, t)


def gerar_analise_diaria(records):
    now       = datetime.now()
    dia_num   = now.weekday()  # 0=Segunda...6=Domingo
    DIAS      = ['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']
    PLURAL    = ['Segundas','Terças','Quartas','Quintas','Sextas','Sábados','Domingos']
    dia_semana = DIAS[dia_num]
    plural     = PLURAL[dia_num]

    data_atual = now.strftime('%d/%m/%Y')
    hora_atual = now.strftime('%H:%M')

    hist  = [r for r in records if r['dia'] == dia_semana]
    datas = sorted(set(r['data'] for r in hist))
    nDias = len(datas)

    if nDias == 0:
        return (
            f"📊 *ANÁLISE DE {plural.upper()} — GUARDA MUNICIPAL BC*\n"
            f"📅 {dia_semana}, {data_atual} às {hora_atual}\n\n"
            f"⚠️ Nenhuma ocorrência registrada em {plural} até o momento."
        )

    total = len(hist)
    media = round(total / nDias, 1)

    tipos    = top_n(hist, 'tipo', 6)
    bairros  = top_n(hist, 'bairro', 5)
    ruas     = top_n(hist, 'endereco', 3)
    turnos_c = {t: sum(1 for r in hist if r['turno'] == t) for t in ORDEM_TURNO}
    top_tipo   = tipos[0] if tipos else ('–', 0)
    top_bairro = bairros[0] if bairros else ('–', 0)
    top_turno  = max(turnos_c, key=turnos_c.get)

    datas_fmt = ', '.join(d[8:10]+'/'+d[5:7] for d in datas[-5:])

    linhas = [
        f"📊 *ANÁLISE DE {plural.upper()} — GUARDA MUNICIPAL BC*",
        f"📅 {dia_semana}, {data_atual} às {hora_atual}",
        "",
        f"*📋 RESUMO HISTÓRICO DE {plural.upper()}*",
        f"Total de ocorrências: *{total}*",
        f"{plural} com dados: *{nDias}* | Média: *{media}* oc./dia",
        f"Tipo mais frequente: *{top_tipo[0]}* ({top_tipo[1]} — {pct(top_tipo[1], total)})",
        f"Bairro mais afetado: *{top_bairro[0]}* ({top_bairro[1]} oc.)",
        f"Turno crítico: *{top_turno}* ({turnos_c[top_turno]} oc. — {pct(turnos_c[top_turno], total)})",
        "",
        f"*🔴 TIPOS EM {plural.upper()}*",
        *[f"• {e[0]}: {e[1]} total ({pct(e[1], total)}) | média {round(e[1]/nDias,1)}/dia"
          for e in tipos],
        "",
        f"*📍 BAIRROS EM {plural.upper()}*",
        *[f"{i+1}. {e[0]}: {e[1]} oc. ({pct(e[1], total)})"
          for i, e in enumerate(bairros)],
        "",
        f"*⏰ TURNOS EM {plural.upper()}*",
        *[f"• {t}: {turnos_c[t]} ({pct(turnos_c[t], total)}) | média {round(turnos_c[t]/nDias,1)}/dia"
          for t in ORDEM_TURNO],
        "",
        f"*🎯 RECOMENDAÇÕES*",
    ]

    recs = []
    if bairros:
        recs.append(f"Reforçar guarnição no {bairros[0][0]} — bairro mais afetado em {plural}.")
    if ruas:
        recs.append(f"Atenção especial a {ruas[0][0]} — logradouro de maior risco.")
    recs.append(f"Turno crítico: {top_turno} concentra {pct(turnos_c[top_turno], total)} das ocorrências.")
    if tipos:
        recs.append(f"Crime mais frequente: {tipos[0][0]}. Orientar guarnições para abordagem preventiva.")

    linhas += [f"{i+1}. {r}" for i, r in enumerate(recs)]
    linhas += [
        "",
        f"_{plural} analisados: {datas_fmt}_",
        "_Guarda Municipal de Balneário Camboriú_",
        "_Secretaria de Segurança e Ordem Pública_",
    ]

    return '\n'.join(linhas)


def gerar_previsao(records):
    now      = datetime.now()
    dia_num  = now.weekday()
    DIAS     = ['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']
    PLURAL   = ['Segundas','Terças','Quartas','Quintas','Sextas','Sábados','Domingos']
    dia_semana = DIAS[dia_num]
    plural     = PLURAL[dia_num]

    data_atual = now.strftime('%d/%m/%Y')
    hora_atual = now.strftime('%H:%M')
    t_atual    = turno_atual()
    t_proximo  = proxima_turno(t_atual)

    hist  = [r for r in records if r['dia'] == dia_semana]
    datas = sorted(set(r['data'] for r in hist))
    nDias = len(datas)

    if nDias == 0:
        return (
            f"📈 *PREVISÃO DE {plural.upper()} — GUARDA MUNICIPAL BC*\n"
            f"📅 {dia_semana}, {data_atual} às {hora_atual}\n\n"
            f"⚠️ Nenhuma ocorrência registrada em {plural} até o momento."
        )

    per_day = [len([r for r in hist if r['data'] == d]) for d in datas]
    mean    = sum(per_day) / nDias
    std_dev = math.sqrt(sum((n - mean)**2 for n in per_day) / nDias)
    min_exp = max(0, round(mean - std_dev))
    max_exp = round(mean + std_dev)

    turnos_c  = {t: sum(1 for r in hist if r['turno'] == t) for t in ORDEM_TURNO}
    total_t   = len(hist) or 1
    top_turno = max(turnos_c, key=turnos_c.get)

    tipos   = top_n(hist, 'tipo', 5)
    bairros = top_n(hist, 'bairro', 5)
    ruas    = top_n(hist, 'endereco', 2)
    max_b   = bairros[0][1] if bairros else 1

    # Tendência
    trend_str = 'Estável'
    trend_icon = '→'
    trend_diff = 0.0
    if len(datas) >= 4:
        mid  = len(datas) // 2
        avg1 = sum(len([r for r in hist if r['data'] == d]) for d in datas[:mid]) / mid
        avg2 = sum(len([r for r in hist if r['data'] == d]) for d in datas[mid:]) / (len(datas) - mid)
        trend_diff = avg2 - avg1
        if trend_diff > 1.5:
            trend_str, trend_icon = 'Crescente', '↗'
        elif trend_diff < -1.5:
            trend_str, trend_icon = 'Decrescente', '↘'

    # Risco geral
    risk_score = 0
    if mean > 8: risk_score += 2
    elif mean > 4: risk_score += 1
    if trend_icon == '↗': risk_score += 2
    if turnos_c[t_atual] / total_t > 0.35: risk_score += 1
    risk_level = 'ALTO' if risk_score >= 4 else 'MÉDIO' if risk_score >= 2 else 'BAIXO'
    risk_emoji = '🔴' if risk_level == 'ALTO' else '🟡' if risk_level == 'MÉDIO' else '🟢'

    linhas = [
        f"📈 *PREVISÃO DE RISCO — {plural.upper()} — GUARDA MUNICIPAL BC*",
        f"📅 {dia_semana}, {data_atual} | Gerado às {hora_atual}",
        "",
        f"*{risk_emoji} NÍVEL DE RISCO GERAL: {risk_level}*",
        f"Previsão: *{min_exp}–{max_exp}* ocorrências esperadas",
        f"Tendência: *{trend_str} {trend_icon}*",
        f"Base histórica: *{nDias}* {plural} analisados | Média: *{round(mean,1)}* oc./dia",
        "",
        f"*⏰ RISCO POR TURNO*",
    ]

    for t in ORDEM_TURNO:
        n   = turnos_c[t]
        p   = pct(n, total_t)
        mark = ' ◀ AGORA' if t == t_atual else ' ⏳ Próximo' if t == t_proximo else ''
        linhas.append(f"• {t}: {n} oc. ({p}){mark}")

    linhas += ["", "*📍 BAIRROS EM ALERTA*"]
    for i, (nome, qtd) in enumerate(bairros):
        score = qtd / max_b
        nivel = '🔴 ALTO' if score > 0.6 else '🟡 MÉDIO' if score > 0.3 else '🟢 BAIXO'
        linhas.append(f"{i+1}. {nome}: {qtd} oc. — {nivel}")

    linhas += ["", "*🔴 CRIMES MAIS PROVÁVEIS*"]
    for i, (tipo, qtd) in enumerate(tipos):
        linhas.append(f"{i+1}. {tipo}: {qtd} ({pct(qtd, len(hist))})")

    linhas += ["", "*📊 HISTÓRICO POR DIA*"]
    for d in datas:
        n   = len([r for r in hist if r['data'] == d])
        fmt = d[8:10] + '/' + d[5:7]
        st  = 'ACIMA' if n > mean + std_dev else 'ABAIXO' if n < mean - std_dev else 'NORMAL'
        linhas.append(f"• {fmt}: {n} oc. — {st}")

    # Orientações
    ori = []
    if bairros:
        ori.append(f"Reforçar guarnição no {bairros[0][0]} — bairro historicamente mais afetado em {plural}.")
    if ruas:
        ori.append(f"Estabelecer ronda em {ruas[0][0]} — logradouro de maior risco em {plural}.")
    turno_pcts = {
        'Madrugada': f"Atenção ao turno da Madrugada (00h–05h) — concentra {pct(turnos_c['Madrugada'], total_t)} das ocorrências em {plural}.",
        'Manhã':     f"Turno Matutino (06h–11h) é crítico em {plural} ({pct(turnos_c['Manhã'], total_t)}). Ampliar presença.",
        'Tarde':     f"Período Vespertino (12h–17h) de alto risco em {plural} ({pct(turnos_c['Tarde'], total_t)}). Atenção à movimentação.",
        'Noite':     f"Turno Noturno (18h–23h) concentra {pct(turnos_c['Noite'], total_t)} dos crimes em {plural}. Coordenar blitz.",
    }
    ori.append(turno_pcts[top_turno])
    if tipos:
        ori.append(f"Crime mais frequente: {tipos[0][0]}. Orientar guarnições com abordagem preventiva.")
    if trend_icon == '↗':
        ori.append(f"⚠️ Tendência crescente detectada em {plural} (+{round(trend_diff,1)} oc./semana). Considerar reforço de efetivo.")
    ori.append("Compartilhar esta previsão com todas as guarnições antes do início do turno.")

    linhas += ["", "*🎯 ORIENTAÇÕES PREVENTIVAS*"]
    linhas += [f"{i+1}. {o}" for i, o in enumerate(ori)]
    linhas += [
        "",
        "_Guarda Municipal de Balneário Camboriú_",
        "_Secretaria de Segurança e Ordem Pública_",
    ]

    return '\n'.join(linhas)


def enviar_telegram(token, chat_id, texto):
    url  = f"https://api.telegram.org/bot{token}/sendMessage"
    data = json.dumps({
        "chat_id":    chat_id,
        "text":       texto,
        "parse_mode": "Markdown",
    }).encode('utf-8')
    req = urllib.request.Request(url, data=data,
                                 headers={'Content-Type': 'application/json'})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read().decode())
            if result.get('ok'):
                print("  OK - Mensagem enviada com sucesso.")
            else:
                print(f"  ERRO Telegram: {result.get('description')}")
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print(f"  ERRO HTTP {e.code}: {body}")
    except Exception as e:
        print(f"  ERRO ao enviar: {e}")


def get_chat_id(token):
    url = f"https://api.telegram.org/bot{token}/getUpdates"
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            data = json.loads(resp.read().decode())
        updates = data.get('result', [])
        if not updates:
            print("\nNenhuma mensagem recebida ainda.")
            print("  -> Envie qualquer mensagem no grupo onde o bot esta")
            print("  -> Depois rode novamente: python enviar_relatorio.py --get-chat-id")
            return
        seen = set()
        print("\nChat IDs encontrados:")
        for u in updates:
            msg = u.get('message') or u.get('channel_post') or {}
            chat = msg.get('chat', {})
            cid  = chat.get('id')
            name = chat.get('title') or chat.get('first_name') or chat.get('username') or '?'
            if cid and cid not in seen:
                seen.add(cid)
                tipo = chat.get('type', '?')
                print(f"  ID: {cid}  |  Nome: {name}  |  Tipo: {tipo}")
        print("\nCole o ID desejado em CHAT_ID no topo do arquivo enviar_relatorio.py")
    except Exception as e:
        print(f"Erro: {e}")


def validar_config():
    if BOT_TOKEN == "SEU_TOKEN_AQUI":
        print("ERRO: Configure BOT_TOKEN no topo do arquivo enviar_relatorio.py")
        sys.exit(1)
    if CHAT_ID == "SEU_CHAT_ID_AQUI":
        print("ERRO: Configure CHAT_ID no topo do arquivo enviar_relatorio.py")
        sys.exit(1)


if __name__ == '__main__':
    if '--get-chat-id' in sys.argv:
        if BOT_TOKEN == "SEU_TOKEN_AQUI":
            print("ERRO: Configure BOT_TOKEN primeiro.")
            sys.exit(1)
        get_chat_id(BOT_TOKEN)
        sys.exit(0)

    validar_config()

    print("Carregando dados da planilha...")
    records = carregar_dados()
    print(f"  {len(records)} registros carregados.")

    print("\nGerando Analise Diaria...")
    texto_analise = gerar_analise_diaria(records)

    print("Gerando Previsao de Risco...")
    texto_previsao = gerar_previsao(records)

    print("\nEnviando para o Telegram...")
    print("  -> Analise Diaria...")
    enviar_telegram(BOT_TOKEN, CHAT_ID, texto_analise)

    print("  -> Previsao de Risco...")
    enviar_telegram(BOT_TOKEN, CHAT_ID, texto_previsao)

    print("\nConcluido.")
