#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extrai dados de Producao Ambulatorial (SIA) do TabNet Rio de Janeiro para
as Policlinicas/Centros Cariocas indicados, por Profissional-CBO x Mes de
Atendimento, para Quantidade Apresentada e Quantidade Aprovada, de 2023 em
diante. Resultado salvo em CSV "long format" (checkpoint) + consolidado depois
em .xlsx por outro script.
"""
import concurrent.futures
import csv
import html
import os
import re
import sys
import time
import urllib.parse
import urllib.request

BASE_URL = "https://tabnet.rio.rj.gov.br/cgi-bin/tabnet?sia/definicoes/producao_2008.def"

ESTABELECIMENTOS = {
    "717": "SMS POLICLINICA ANTONIO RIBEIRO NETTO AP 10",
    "718": "SMS POLICLINICA CARLOS ALBERTO NASCIMENTO AP 52",
    "719": "SMS POLICLINICA HELIO PELLEGRINO AP 22",
    "720": "SMS POLICLINICA JOSE PARANHOS FONTENELLE AP 31",
    "721": "SMS POLICLINICA LINCOLN DE FREITAS FILHO AP 53",
    "722": "SMS POLICLINICA MANOEL GUILHERME PAM BANGU AP 51",
    "723": "SMS POLICLINICA NEWTON ALVES CARDOZO AP 31",
    "724": "SMS POLICLINICA NEWTON BETHLEM AP 40",
    "725": "SMS POLICLINICA ROCHA MAIA AP 21",
    "726": "SMS POLICLINICA RODOLPHO ROCCO AP 32",
    "327": "SMS CENTRO CARIOCA DE ESPECIALIDADE DA ZONA OESTE AP 52",
    "330": "SMS CENTRO CARIOCA DE REABILITACAO DA ZONA OESTE AP 52",
    "329": "SMS CENTRO CARIOCA DE HEMODIALISE AP 52",
}

MESES_PT = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}

def gerar_arquivos(ano_ini, mes_ini, ano_fim, mes_fim):
    files = []
    y, m = ano_ini, mes_ini
    while (y, m) <= (ano_fim, mes_fim):
        yy = str(y)[2:]
        mm = f"{m:02d}"
        fname = f"parj{yy}{mm}.dbf"
        label = f"{MESES_PT[m]}/{y}"
        files.append((fname, label, y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return files

TODOS_ARQUIVOS = gerar_arquivos(2023, 1, 2026, 7)

CONTEUDOS = {
    "apresentada": "Qtd.Apresentada",
    "aprovada": "Qtd.Aprovada",
}

DEFAULT_FILTERS = [
    ("SAREAS_RJ_(caps)", "TODAS_AS_CATEGORIAS__"),
    ("SGerência/Adm(Terceiro)", "TODAS_AS_CATEGORIAS__"),
    ("SREGULACAO", "TODAS_AS_CATEGORIAS__"),
    ("SEstabel-CNES-RJ", "TODAS_AS_CATEGORIAS__"),
    ("STipo_Unidade_SMS", "TODAS_AS_CATEGORIAS__"),
    ("STipo_Estabelecimento_SUS", "TODAS_AS_CATEGORIAS__"),
    ("STipoUnidade_Geral_Agrupada_SMS", "TODAS_AS_CATEGORIAS__"),
    ("SEsfera_Subgeral_SMS", "TODAS_AS_CATEGORIAS__"),
    ("SProd.Aprovada/Não", "TODAS_AS_CATEGORIAS__"),
    ("SInd.Erro_QTDE_produzida", "TODAS_AS_CATEGORIAS__"),
    ("SMês_Cobrança", "TODAS_AS_CATEGORIAS__"),
    ("SQuadr.Cobrança", "TODAS_AS_CATEGORIAS__"),
    ("SQuadr.Atendimento", "TODAS_AS_CATEGORIAS__"),
    ("SGRUPO", "TODAS_AS_CATEGORIAS__"),
    ("SSUBGRUPO", "TODAS_AS_CATEGORIAS__"),
    ("SFORMA_ORGANIZAÇÃO", "TODAS_AS_CATEGORIAS__"),
    ("SPROCEDIMENTO", "TODAS_AS_CATEGORIAS__"),
    ("SCOMPLEXIDADE", "TODAS_AS_CATEGORIAS__"),
    ("SDoc.Origem", "TODAS_AS_CATEGORIAS__"),
    ("STIPO_FINANCIAMENTO", "TODAS_AS_CATEGORIAS__"),
    ("STIPO_Fin/SubtipoFin.", "TODAS_AS_CATEGORIAS__"),
    ("SProfissional-CBO", "TODAS_AS_CATEGORIAS__"),
    ("SEsfera_Administr.", "TODAS_AS_CATEGORIAS__"),
    ("SNATUREZA", "TODAS_AS_CATEGORIAS__"),
    ("SMUNICIPIO_DO_PACIENTE", "TODAS_AS_CATEGORIAS__"),
    ("SRAÇA/COR_PACIENTE", "TODAS_AS_CATEGORIAS__"),
    ("SSEXO__DO_PACIENTE", "TODAS_AS_CATEGORIAS__"),
    ("SIDADE_DO_PACIENTE", "TODAS_AS_CATEGORIAS__"),
]

CHUNK_SIZE = 6


def montar_body(estab_codigo, conteudo_valor, arquivos_chunk):
    fields = [
        ("Linha", "Profissional-CBO"),
        ("Coluna", "Mês_Atendimento"),
        ("Incremento", conteudo_valor),
    ]
    fields += [("Arquivos", f[0]) for f in arquivos_chunk]
    fields += DEFAULT_FILTERS
    fields += [("SEstabel-NOME-RJ", estab_codigo)]
    fields += [("formato", "table"), ("mostre", "Mostra")]
    return urllib.parse.urlencode(fields, encoding="latin-1").encode("latin-1")


ROW_RE = re.compile(r"<TR[^>]*>(.*?)</TR>", re.IGNORECASE | re.DOTALL)
TD_RE = re.compile(r"<T[DH][^>]*>(.*?)(?=<T[DH][^>]*>|</TR>|\Z)", re.IGNORECASE | re.DOTALL)
TAG_RE = re.compile(r"<[^>]+>")


def limpar_texto(raw):
    t = TAG_RE.sub("", raw)
    t = html.unescape(t)
    return t.strip()


def parse_number(txt):
    txt = txt.strip()
    if txt in ("-", "", "..", "0"):
        return 0.0 if txt != "0" else 0.0
    txt = txt.replace(".", "").replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return None


ROW_SPLIT_RE = re.compile(r"<TR[^>]*>", re.IGNORECASE)
CELL_SPLIT_RE = re.compile(r"<T[DH][^>]*>", re.IGNORECASE)


def _split_row_cells(row_text):
    parts = CELL_SPLIT_RE.split(row_text)[1:]
    return [limpar_texto(p) for p in parts]


def parse_tabela(html_text):
    """Retorna (col_labels, linhas) onde linhas = [(rotulo_linha, [valores por coluna])]
    Ignora a linha TOTAL e a coluna Total.
    O HTML do TabNet nao fecha as tags <TR>/<TD>/<TBODY> corretamente, entao o
    parsing e feito por divisao de texto (split) em vez de regex com fechamento."""
    marker = html_text.upper().find('CLASS="TABDADOS"')
    if marker == -1:
        return None, []
    table_start = html_text.upper().rfind("<TABLE", 0, marker)
    if table_start == -1:
        return None, []
    table_end = html_text.upper().find("</TABLE>", table_start)
    if table_end == -1:
        table_end = len(html_text)

    thead_start = html_text.upper().find("<THEAD", table_start, table_end)
    tfoot_start = html_text.upper().find("<TFOOT", thead_start, table_end)
    tbody_start = html_text.upper().find("<TBODY", tfoot_start if tfoot_start != -1 else thead_start, table_end)
    header_region_end = tfoot_start if tfoot_start != -1 else (tbody_start if tbody_start != -1 else table_end)
    if thead_start == -1:
        return None, []
    header_region = html_text[thead_start:header_region_end]

    header_marker = header_region.upper().find('TR VALIGN="BOTTOM"')
    if header_marker == -1:
        return None, []
    gt_idx = header_region.find(">", header_marker)
    header_row_text = header_region[gt_idx + 1 :]
    header_cells = _split_row_cells(header_row_text)
    if not header_cells:
        return None, []
    col_labels = header_cells[1:-1] if len(header_cells) > 2 else header_cells[1:]

    if tbody_start == -1:
        return col_labels, []
    body_region = html_text[tbody_start:table_end]
    gt_idx = body_region.find(">")
    body_region = body_region[gt_idx + 1 :]

    linhas = []
    for row_text in ROW_SPLIT_RE.split(body_region):
        if not row_text.strip():
            continue
        textos = _split_row_cells(row_text)
        if not textos or not textos[0]:
            continue
        rotulo = textos[0]
        if rotulo.upper() == "TOTAL":
            continue
        valores_txt = textos[1:]
        if len(valores_txt) > len(col_labels):
            valores_txt = valores_txt[: len(col_labels)]
        valores = [parse_number(v) for v in valores_txt]
        linhas.append((rotulo, valores))
    return col_labels, linhas


def fetch(estab_codigo, conteudo_valor, arquivos_chunk, tentativas=6):
    body = montar_body(estab_codigo, conteudo_valor, arquivos_chunk)
    for tentativa in range(1, tentativas + 1):
        try:
            req = urllib.request.Request(
                BASE_URL,
                data=body,
                headers={
                    "Content-Type": "application/x-www-form-urlencoded",
                    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36",
                    "Connection": "close",
                },
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=100) as resp:
                raw = resp.read()
            return raw.decode("latin-1", errors="replace")
        except Exception as e:
            if tentativa == tentativas:
                print(f"  [ERRO] estab={estab_codigo} conteudo={conteudo_valor} chunk={arquivos_chunk[0][1]}..{arquivos_chunk[-1][1]}: {e}", file=sys.stderr, flush=True)
                return None
            espera = min(8 * tentativa, 45)
            time.sleep(espera)
    return None


def chunked(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i : i + n]


def task_key(tarefa):
    estab_codigo, conteudo_nome, _, arquivos_chunk = tarefa
    return f"{estab_codigo}|{conteudo_nome}|{arquivos_chunk[0][0]}"


def processar_tarefa(tarefa):
    estab_codigo, conteudo_nome, conteudo_valor, arquivos_chunk = tarefa
    html_text = fetch(estab_codigo, conteudo_valor, arquivos_chunk)
    resultados = []
    if html_text is None:
        return False, resultados
    col_labels, linhas = parse_tabela(html_text)
    if col_labels is None:
        return False, resultados
    for rotulo, valores in linhas:
        for col_label, valor in zip(col_labels, valores):
            if valor is None:
                continue
            resultados.append(
                {
                    "estabelecimento_codigo": estab_codigo,
                    "estabelecimento_nome": ESTABELECIMENTOS[estab_codigo],
                    "conteudo": conteudo_nome,
                    "profissional_cbo": rotulo,
                    "mes_atendimento": col_label,
                    "valor": valor,
                }
            )
    return True, resultados


def main():
    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
    os.makedirs(out_dir, exist_ok=True)
    out_csv = os.path.join(out_dir, "tabnet_raw_long.csv")
    checkpoint_path = os.path.join(out_dir, "tabnet_checkpoint.txt")

    tarefas = []
    for estab_codigo in ESTABELECIMENTOS:
        for conteudo_nome, conteudo_valor in CONTEUDOS.items():
            for chunk in chunked(TODOS_ARQUIVOS, CHUNK_SIZE):
                tarefas.append((estab_codigo, conteudo_nome, conteudo_valor, chunk))

    concluidas_antes = set()
    if os.path.exists(checkpoint_path):
        with open(checkpoint_path, encoding="utf-8") as f:
            concluidas_antes = set(line.strip() for line in f if line.strip())

    pendentes = [t for t in tarefas if task_key(t) not in concluidas_antes]

    print(f"Total de tarefas: {len(tarefas)} | ja concluidas: {len(concluidas_antes)} | pendentes: {len(pendentes)}", flush=True)

    fieldnames = ["estabelecimento_codigo", "estabelecimento_nome", "conteudo", "profissional_cbo", "mes_atendimento", "valor"]
    novo_csv = not os.path.exists(out_csv)
    concluidas = 0
    falhas = []
    t0 = time.time()
    with open(out_csv, "a", newline="", encoding="utf-8") as f, open(checkpoint_path, "a", encoding="utf-8") as cp:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if novo_csv:
            writer.writeheader()
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as ex:
            futures = {}
            for t in pendentes:
                futures[ex.submit(processar_tarefa, t)] = t
                time.sleep(0.3)
            for fut in concurrent.futures.as_completed(futures):
                tarefa = futures[fut]
                estab_codigo, conteudo_nome, _, chunk = tarefa
                try:
                    ok, resultados = fut.result()
                except Exception as e:
                    ok, resultados = False, []
                    print(f"  [FALHA] {tarefa[:2]} chunk={chunk[0][1]}: {e}", file=sys.stderr, flush=True)
                for r in resultados:
                    writer.writerow(r)
                f.flush()
                if ok:
                    cp.write(task_key(tarefa) + "\n")
                    cp.flush()
                else:
                    falhas.append(tarefa)
                concluidas += 1
                elapsed = time.time() - t0
                status = "OK" if ok else "FALHOU"
                print(f"[{concluidas}/{len(pendentes)}] estab={estab_codigo} conteudo={conteudo_nome} chunk={chunk[0][1]}-{chunk[-1][1]} linhas={len(resultados)} {status} ({elapsed:.0f}s)", flush=True)

    print(f"Concluido em {time.time()-t0:.0f}s. CSV salvo em: {out_csv}", flush=True)
    if falhas:
        print(f"ATENCAO: {len(falhas)} chunk(s) falharam apos todas as tentativas. Rode o script novamente para retentar (ele retoma automaticamente).", flush=True)


if __name__ == "__main__":
    main()
