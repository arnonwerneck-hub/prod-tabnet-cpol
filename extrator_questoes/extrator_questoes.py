# -*- coding: utf-8 -*-
"""
=====================================================================
 EXTRATOR DE QUESTÕES DE CONCURSOS PÚBLICOS (PDF -> Banco de Dados)
=====================================================================

O QUE ESTE PROGRAMA FAZ (em linguagem simples):

  1. Olha a pasta "pdfs" (que fica ao lado deste arquivo) e encontra
     todos os arquivos PDF que estão lá.
  2. Para cada PDF que AINDA NÃO foi lido antes, ele:
       - lê o texto de todas as páginas;
       - tenta descobrir Banca, Ano, Órgão e Cargo pelo cabeçalho;
       - separa as questões (enunciado + alternativas A, B, C, D, E);
       - tenta descobrir a Disciplina e o Assunto de cada questão;
       - procura o gabarito (no fim da prova ou num PDF de gabarito);
       - recorta as imagens das questões e salva em JPG pequeno.
  3. Guarda tudo num único arquivo de banco de dados chamado
     "banco_questoes.db" (formato SQLite: um arquivo só, leve).
  4. Ao final, gera também uma cópia em JSON e em CSV (abre no Excel)
     para você consultar com facilidade.

SINCRONIZAÇÃO (não duplicar):
  O programa calcula uma "impressão digital" (hash) de cada PDF.
  Se você rodar o programa de novo, ele PULA os PDFs que já leu e
  processa só os novos. Mesmo que você renomeie um PDF já lido, ele
  reconhece que é o mesmo arquivo. E se a mesma questão aparecer em
  dois PDFs diferentes, ela só é gravada uma vez.

COMO RODAR:
  Abra o terminal nesta pasta e digite:
      python extrator_questoes.py

  (Veja o arquivo LEIAME.md para o passo a passo completo.)

Linhas que começam com "#" são COMENTÁRIOS: o Python ignora, elas
existem só para explicar o código para você.
"""

# ---------------------------------------------------------------------
# IMPORTAÇÕES
# "import" = trazer ferramentas prontas para dentro do programa.
# ---------------------------------------------------------------------
import csv            # para gerar o arquivo .csv (planilha)
import hashlib        # para calcular a "impressão digital" dos arquivos
import json           # para gravar dados no formato JSON
import re             # "Expressões Regulares" (Regex): busca de padrões no texto
import sqlite3        # banco de dados SQLite (já vem com o Python)
import sys            # informações do sistema (usado para mensagens de erro)
import unicodedata    # para remover acentos ao comparar palavras
from collections import Counter   # conta quantas vezes algo aparece
from datetime import datetime     # data e hora do processamento
from pathlib import Path          # manipular caminhos de pastas/arquivos

# A biblioteca "pdfplumber" é a única que precisa ser instalada à parte.
# Se ela não estiver instalada, mostramos uma mensagem amigável.
try:
    import pdfplumber
except ImportError:
    print("\n[ERRO] A biblioteca 'pdfplumber' não está instalada.")
    print("Rode no terminal:   pip install pdfplumber\n")
    sys.exit(1)


# =====================================================================
#  CONFIGURAÇÕES  (você pode mudar os valores abaixo se quiser)
# =====================================================================

# Pasta onde este arquivo .py está guardado.
PASTA_DO_PROGRAMA = Path(__file__).resolve().parent

# Pasta onde você deve colocar os PDFs das provas (e dos gabaritos).
PASTA_PDFS = PASTA_DO_PROGRAMA / "pdfs"

# Pasta onde as imagens recortadas das questões serão salvas.
PASTA_IMAGENS = PASTA_DO_PROGRAMA / "imagens"

# Arquivo do banco de dados (tudo fica guardado aqui).
ARQUIVO_BANCO = PASTA_DO_PROGRAMA / "banco_questoes.db"

# Cópias para consulta (são recriadas toda vez que o programa roda).
ARQUIVO_JSON = PASTA_DO_PROGRAMA / "banco_questoes.json"
ARQUIVO_CSV = PASTA_DO_PROGRAMA / "banco_questoes.csv"

# Texto usado quando o programa não consegue descobrir uma informação.
NAO_IDENTIFICADO = "Não Identificado"

# Configurações das imagens (para manter tudo leve):
SALVAR_IMAGENS = True          # mude para False se não quiser imagens
IMAGEM_LARGURA_MAXIMA = 900    # em pixels; imagens maiores são reduzidas
IMAGEM_QUALIDADE_JPG = 70      # de 1 a 95; menor = arquivo menor
IMAGEM_TAMANHO_MINIMO = 40     # em pontos; ignora ícones minúsculos
IMAGEM_RESOLUCAO = 110         # "DPI" do recorte (qualidade da leitura)

# ---------------------------------------------------------------------
# LISTA DE BANCAS CONHECIDAS
# À esquerda: o nome "oficial" que será gravado.
# À direita: as formas como o nome pode aparecer no PDF.
# Para adicionar uma banca, copie uma linha e altere.
# ---------------------------------------------------------------------
BANCAS = {
    "Cebraspe": ["CEBRASPE", "CESPE", "CENTRO BRASILEIRO DE PESQUISA EM AVALIACAO"],
    "FGV": ["FGV", "FUNDACAO GETULIO VARGAS"],
    "FCC": ["FCC", "FUNDACAO CARLOS CHAGAS"],
    "Vunesp": ["VUNESP"],
    "Cesgranrio": ["CESGRANRIO"],
    "IBFC": ["IBFC"],
    "Quadrix": ["QUADRIX"],
    "Idecan": ["IDECAN"],
    "Instituto AOCP": ["AOCP"],
    "Consulplan": ["CONSULPLAN"],
    "Fundatec": ["FUNDATEC"],
    "Iades": ["IADES"],
    "Ibade": ["IBADE"],
    "Selecon": ["SELECON"],
    "Fepese": ["FEPESE"],
    "Fumarc": ["FUMARC"],
    "Objetiva": ["OBJETIVA CONCURSOS"],
    "Instituto Access": ["INSTITUTO ACCESS"],
    "Avança SP": ["AVANCA SP", "AVANCASP"],
    "Instituto Consulplan": ["INSTITUTO CONSULPLAN"],
    "Cetro": ["CETRO"],
    "Esaf": ["ESAF"],
    "Comperve": ["COMPERVE"],
    "Funrio": ["FUNRIO"],
    "Ibam": ["IBAM"],
    "Ibest": ["IBEST"],
    "Instituto Mais": ["INSTITUTO MAIS"],
    "Legalle": ["LEGALLE"],
    "Nucepe": ["NUCEPE"],
    "UFPR / NC": ["NUCLEO DE CONCURSOS"],
}

# ---------------------------------------------------------------------
# PALAVRAS QUE INDICAM O NOME DO ÓRGÃO
# Se uma linha do cabeçalho começar (ou contiver) uma destas
# expressões, consideramos que ela é o nome do órgão.
# ---------------------------------------------------------------------
PALAVRAS_ORGAO = [
    "PREFEITURA", "TRIBUNAL", "SECRETARIA", "MINISTERIO", "POLICIA",
    "CAMARA MUNICIPAL", "CAMARA DOS DEPUTADOS", "ASSEMBLEIA LEGISLATIVA",
    "SENADO", "BANCO DO BRASIL", "BANCO CENTRAL", "CAIXA ECONOMICA",
    "UNIVERSIDADE", "INSTITUTO FEDERAL", "INSTITUTO NACIONAL",
    "DEFENSORIA", "PROCURADORIA", "MINISTERIO PUBLICO", "AGENCIA NACIONAL",
    "CONSELHO REGIONAL", "CONSELHO FEDERAL", "CORPO DE BOMBEIROS",
    "DEPARTAMENTO", "RECEITA FEDERAL", "AUTARQUIA", "EMPRESA BRASILEIRA",
    "PETROBRAS", "CORREIOS", "GOVERNO DO ESTADO", "ESTADO D", "MUNICIPIO D",
    "COMPANHIA", "SERVICO AUTONOMO", "FUNDACAO HOSPITALAR", "HOSPITAL",
    "CONTROLADORIA", "AGENCIA ESTADUAL", "DETRAN", "INSS", "IBGE",
]

# ---------------------------------------------------------------------
# PALAVRAS QUE INDICAM O CARGO
# ---------------------------------------------------------------------
PALAVRAS_CARGO = [
    "ANALISTA", "TECNICO", "AUDITOR", "AGENTE", "ESCRIVAO", "DELEGADO",
    "PROFESSOR", "ASSISTENTE", "OFICIAL", "INSPETOR", "PERITO", "FISCAL",
    "ESCRITURARIO", "AUXILIAR", "PROCURADOR", "DEFENSOR", "JUIZ",
    "PROMOTOR", "ENFERMEIRO", "MEDICO", "ENGENHEIRO", "CONTADOR",
    "GUARDA", "SOLDADO", "INVESTIGADOR", "PAPILOSCOPISTA", "ADMINISTRADOR",
    "ESPECIALISTA", "GESTOR", "ASSESSOR", "ADVOGADO", "BIBLIOTECARIO",
    "PSICOLOGO", "ASSISTENTE SOCIAL", "NUTRICIONISTA", "FARMACEUTICO",
    "ODONTOLOGO", "CIRURGIAO", "MOTORISTA", "OPERADOR", "ARQUITETO",
]

# ---------------------------------------------------------------------
# DISCIPLINAS (MATÉRIAS)
# À esquerda: nome que será gravado.
# À direita: como o título da seção pode aparecer na prova.
# O programa também usa "palavras-pista" (mais abaixo) para tentar
# adivinhar a disciplina quando a prova não tem títulos de seção.
# ---------------------------------------------------------------------
DISCIPLINAS = {
    "Língua Portuguesa": ["LINGUA PORTUGUESA", "PORTUGUES", "LINGUAGEM", "INTERPRETACAO DE TEXTO"],
    "Matemática": ["MATEMATICA FINANCEIRA", "MATEMATICA"],
    "Raciocínio Lógico": ["RACIOCINIO LOGICO", "RACIOCINIO LOGICO-MATEMATICO", "RACIOCINIO ANALITICO"],
    "Informática": ["INFORMATICA", "NOCOES DE INFORMATICA", "TECNOLOGIA DA INFORMACAO"],
    "Língua Inglesa": ["LINGUA INGLESA", "INGLES"],
    "Língua Espanhola": ["LINGUA ESPANHOLA", "ESPANHOL"],
    "Direito Constitucional": ["DIREITO CONSTITUCIONAL"],
    "Direito Administrativo": ["DIREITO ADMINISTRATIVO"],
    "Direito Penal": ["DIREITO PENAL"],
    "Direito Processual Penal": ["DIREITO PROCESSUAL PENAL"],
    "Direito Civil": ["DIREITO CIVIL"],
    "Direito Processual Civil": ["DIREITO PROCESSUAL CIVIL"],
    "Direito Tributário": ["DIREITO TRIBUTARIO"],
    "Direito do Trabalho": ["DIREITO DO TRABALHO"],
    "Direito Previdenciário": ["DIREITO PREVIDENCIARIO"],
    "Direitos Humanos": ["DIREITOS HUMANOS"],
    "Legislação": ["LEGISLACAO ESPECIFICA", "LEGISLACAO"],
    "Administração Pública": ["ADMINISTRACAO PUBLICA", "ADMINISTRACAO GERAL"],
    "Contabilidade": ["CONTABILIDADE"],
    "Economia": ["ECONOMIA"],
    "Estatística": ["ESTATISTICA"],
    "Arquivologia": ["ARQUIVOLOGIA"],
    "Atualidades": ["ATUALIDADES", "CONHECIMENTOS GERAIS"],
    "Ética no Serviço Público": ["ETICA NO SERVICO PUBLICO", "ETICA"],
    "Saúde Pública / SUS": ["SAUDE PUBLICA", "LEGISLACAO DO SUS", "POLITICAS DE SAUDE"],
}

# Palavras-pista: se a questão contiver estas palavras, provavelmente
# é desta disciplina. (Usado só quando a prova não tem título de seção.)
PISTAS_DISCIPLINA = {
    "Direito Constitucional": ["constituicao federal", "cf/88", "direitos fundamentais", "emenda constitucional"],
    "Direito Administrativo": ["ato administrativo", "licitacao", "lei 14.133", "lei n 8.112", "servidor publico", "improbidade"],
    "Direito Penal": ["codigo penal", "crime", "pena de reclusao", "dolo", "culpa"],
    "Direito Civil": ["codigo civil", "contrato", "obrigacoes", "posse", "usucapiao"],
    "Informática": ["windows", "excel", "word", "navegador", "linux", "planilha", "e-mail", "hardware", "software"],
    "Raciocínio Lógico": ["proposicao", "tabela-verdade", "silogismo", "conjunto", "probabilidade"],
    "Matemática": ["porcentagem", "juros", "equacao", "fracao", "razao e proporcao"],
    "Língua Portuguesa": ["no texto", "do texto", "concordancia", "crase", "virgula", "regencia", "sentido do texto"],
    "Contabilidade": ["balanco patrimonial", "ativo circulante", "debito", "credito", "demonstracao do resultado"],
}

# ASSUNTOS (tópicos): palavras que, se aparecerem na questão,
# indicam um assunto específico. Organizado por disciplina.
# Fique à vontade para acrescentar os seus.
ASSUNTOS = {
    "Língua Portuguesa": {
        "Crase": ["crase", "acento grave"],
        "Concordância": ["concordancia"],
        "Regência": ["regencia"],
        "Pontuação": ["virgula", "pontuacao", "ponto e virgula"],
        "Colocação Pronominal": ["colocacao pronominal", "proclise", "enclise", "mesoclise"],
        "Interpretação de Texto": ["de acordo com o texto", "segundo o texto", "infere-se", "o autor"],
        "Classes de Palavras": ["substantivo", "adjetivo", "adverbio", "pronome", "preposicao", "conjuncao"],
        "Ortografia e Acentuação": ["ortografia", "acentuacao", "acentuada"],
        "Semântica / Sinônimos": ["sinonimo", "antonimo", "sentido da palavra", "substituido sem prejuizo"],
    },
    "Informática": {
        "Planilhas (Excel/Calc)": ["excel", "planilha", "calc"],
        "Editores de Texto (Word/Writer)": ["word", "writer"],
        "Sistemas Operacionais": ["windows", "linux"],
        "Internet e Navegadores": ["navegador", "chrome", "firefox", "edge", "internet"],
        "Segurança da Informação": ["virus", "malware", "phishing", "firewall", "backup"],
        "Correio Eletrônico": ["e-mail", "outlook", "correio eletronico"],
    },
    "Raciocínio Lógico": {
        "Lógica Proposicional": ["proposicao", "tabela-verdade", "conectivo", "negacao"],
        "Probabilidade": ["probabilidade"],
        "Análise Combinatória": ["combinacao", "arranjo", "permutacao", "anagrama"],
        "Conjuntos": ["conjunto"],
        "Sequências": ["sequencia"],
    },
    "Matemática": {
        "Porcentagem": ["porcentagem", "%"],
        "Juros": ["juros"],
        "Razão e Proporção": ["razao", "proporcao", "regra de tres"],
        "Equações": ["equacao"],
    },
    "Direito Constitucional": {
        "Direitos e Garantias Fundamentais": ["direitos fundamentais", "art. 5", "garantias fundamentais"],
        "Organização do Estado": ["uniao, estados", "competencia da uniao", "municipios"],
        "Poder Legislativo": ["congresso nacional", "processo legislativo"],
        "Poder Judiciário": ["supremo tribunal", "poder judiciario"],
        "Controle de Constitucionalidade": ["controle de constitucionalidade", "adi", "adpf"],
    },
    "Direito Administrativo": {
        "Atos Administrativos": ["ato administrativo"],
        "Licitações e Contratos": ["licitacao", "14.133", "8.666", "pregao"],
        "Agentes Públicos": ["8.112", "servidor publico", "agente publico"],
        "Improbidade Administrativa": ["improbidade"],
        "Princípios da Administração": ["principio da legalidade", "impessoalidade", "moralidade", "eficiencia"],
        "Poderes Administrativos": ["poder de policia", "poder hierarquico", "poder disciplinar"],
    },
}


# =====================================================================
#  EXPRESSÕES REGULARES (REGEX) - os "moldes" de busca no texto
# =====================================================================
# Pense numa Regex como um "molde": o programa procura no texto
# trechos que se encaixem nesse molde.

# Início de questão escrito por extenso: "Questão 12", "QUESTÃO Nº 12"...
RE_QUESTAO_EXTENSO = re.compile(
    r"^\s*QUEST[ÃA]O\s*(?:N[º°o.]?\s*)?(\d{1,3})\b[\s.:\-–—)]*(.*)$",
    re.IGNORECASE,
)

# Início de questão só com número: "12.", "12)", "12 -", "12 Texto..."
# ou até um número sozinho na linha ("12").
RE_QUESTAO_NUMERO = re.compile(r"^\s*(\d{1,3})\s*(?:[.)\-–—:]\s*|\s+(?=\S)|$)(.*)$")

# Alternativa: "A)", "(A)", "a)", "A.", "A -", "A –"
RE_ALTERNATIVA = re.compile(r"^\s*\(?([A-Ea-e])\s*[)\.\-–—]\s*(.*)$")

# Alternativas grudadas na mesma linha: "... (B) texto (C) texto"
RE_ALTERNATIVA_INLINE = re.compile(r"\s\(?([B-Eb-e])\)\s")

# Linhas de "lixo" típico (numeração de página, etc.)
RE_LIXO = [
    re.compile(r"^\s*P[áa]gina\s+\d+(\s+de\s+\d+)?\s*$", re.IGNORECASE),
    re.compile(r"^\s*\d+\s*/\s*\d+\s*$"),            # "3/20"
    re.compile(r"^\s*-\s*\d+\s*-\s*$"),               # "- 3 -"
    re.compile(r"^\s*RASCUNHO\s*$", re.IGNORECASE),
]

# Ano: 4 dígitos começando com 19 ou 20 (ex.: 2023)
RE_ANO = re.compile(r"\b(19[89]\d|20[0-4]\d)\b")

# Cargo escrito com rótulo: "Cargo: Analista Judiciário"
RE_CARGO_ROTULO = re.compile(r"\bCARGO\s*(?:/\s*\w+)?\s*[:\-–]\s*(.+)", re.IGNORECASE)

# Par "número + resposta" de um gabarito: "1-A", "12 C", "3) E", "4: ANULADA"
RE_PAR_GABARITO = re.compile(
    r"\b(\d{1,3})\s*[-–—.:)]?\s*(ANULADA|ANULADO|CERTO|ERRADO|[A-E]|X|\*)(?![A-Za-zÀ-ú])",
    re.IGNORECASE,
)


# =====================================================================
#  FUNÇÕES AUXILIARES (pequenas ferramentas usadas pelo programa)
# =====================================================================

def sem_acento(texto):
    """Remove acentos e deixa em MAIÚSCULAS. Ex.: 'Questão' -> 'QUESTAO'.
    Facilita comparar palavras sem se preocupar com acentuação."""
    texto = unicodedata.normalize("NFD", texto or "")
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto.upper()


def hash_de_arquivo(caminho):
    """Calcula a 'impressão digital' (SHA-256) de um arquivo.
    Dois arquivos idênticos sempre têm a mesma impressão digital,
    mesmo que tenham nomes diferentes."""
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1024 * 1024), b""):
            h.update(bloco)
    return h.hexdigest()


def hash_de_texto(texto):
    """Impressão digital de um texto (usada para não repetir questões)."""
    normal = re.sub(r"\s+", " ", sem_acento(texto)).strip()
    return hashlib.sha1(normal.encode("utf-8")).hexdigest()


def chave_da_prova(nome_arquivo):
    """Cria uma 'chave' a partir do nome do arquivo, usada para ligar
    um PDF de gabarito à sua prova.
    Ex.: 'TRT_2023_Analista_gabarito_definitivo.pdf' -> 'TRT 2023 ANALISTA'
         'TRT_2023_Analista.pdf'                     -> 'TRT 2023 ANALISTA'"""
    nome = sem_acento(Path(nome_arquivo).stem)
    nome = re.sub(r"[^A-Z0-9]+", " ", nome)
    for palavra in ["GABARITOS", "GABARITO", "DEFINITIVO", "PRELIMINAR", "OFICIAL",
                    "PROVA", "CADERNO", "RESPOSTAS"]:
        nome = re.sub(rf"\b{palavra}\b", " ", nome)
    return re.sub(r"\s+", " ", nome).strip()


def juntar_linhas(linhas):
    """Junta várias linhas num único parágrafo, corrigindo palavras
    quebradas com hífen no fim da linha (ex.: 'admi-' + 'nistração')."""
    texto = ""
    for linha in linhas:
        linha = linha.strip()
        if not linha:
            continue
        if texto.endswith("-") and linha[:1].islower():
            texto = texto[:-1] + linha           # junta palavra quebrada
        elif texto:
            texto += " " + linha
        else:
            texto = linha
    return re.sub(r"\s+", " ", texto).strip()


def eh_lixo(linha):
    """Diz se a linha é 'lixo' (número de página, 'Rascunho', etc.)."""
    return any(r.match(linha) for r in RE_LIXO)


# =====================================================================
#  BANCO DE DADOS
# =====================================================================

def abrir_banco():
    """Abre (ou cria, se ainda não existir) o arquivo do banco de dados
    e garante que as 'tabelas' (como abas de uma planilha) existam."""
    con = sqlite3.connect(ARQUIVO_BANCO)
    con.executescript(
        """
        -- Lista dos PDFs já processados (é isso que evita reprocessar)
        CREATE TABLE IF NOT EXISTS arquivos_processados (
            hash_arquivo       TEXT PRIMARY KEY,
            nome_arquivo       TEXT,
            tipo               TEXT,     -- 'prova', 'gabarito' ou 'sem_texto'
            chave_prova        TEXT,
            qtd_questoes       INTEGER,
            data_processamento TEXT
        );

        -- As questões propriamente ditas
        CREATE TABLE IF NOT EXISTS questoes (
            id             TEXT PRIMARY KEY,   -- ID único
            hash_conteudo  TEXT UNIQUE,        -- evita a mesma questão duas vezes
            banca          TEXT,
            ano            TEXT,
            orgao          TEXT,
            cargo          TEXT,
            disciplina     TEXT,
            assunto        TEXT,
            numero         INTEGER,            -- número da questão na prova
            enunciado      TEXT,
            alternativas   TEXT,               -- dicionário {"A": "...", ...} em JSON
            gabarito       TEXT,               -- letra correta
            imagens        TEXT,               -- lista de imagens em JSON
            arquivo_origem TEXT,
            chave_prova    TEXT,
            data_inclusao  TEXT
        );

        -- Respostas lidas de PDFs de gabarito (podem chegar antes da prova)
        CREATE TABLE IF NOT EXISTS gabaritos (
            chave_prova TEXT,
            numero      INTEGER,
            resposta    TEXT,
            definitivo  INTEGER,
            PRIMARY KEY (chave_prova, numero)
        );
        """
    )
    return con


# =====================================================================
#  LEITURA DO PDF
# =====================================================================

def dividir_em_regioes(pagina):
    """Muitas provas são impressas em DUAS COLUNAS. Se lermos a página
    de uma vez, o texto das duas colunas se mistura. Aqui verificamos
    se existe um 'corredor' vazio no meio da página.

    Devolve a lista de 'regiões' da página na ordem certa de leitura:
      - faixas de largura total (ex.: título que atravessa a página);
      - entre elas, a coluna da esquerda e depois a da direita."""
    x0, topo, x1, base = pagina.bbox
    meio = x0 + pagina.width / 2
    palavras = pagina.extract_words()
    if len(palavras) < 15:
        return [pagina.bbox]

    def cruza_meio(p):
        return p["x0"] < meio - 3 and p["x1"] > meio + 3

    # Agrupa as palavras em "linhas visuais" (mesma altura na página).
    linhas_visuais = {}
    for p in palavras:
        linhas_visuais.setdefault(round(p["top"] / 3), []).append(p)

    linhas_que_cruzam = 0
    inicios_direita = Counter()   # onde começa o texto do lado direito
    for ps in linhas_visuais.values():
        if any(cruza_meio(p) for p in ps):
            linhas_que_cruzam += 1
            continue
        direita = [p["x0"] for p in ps if p["x0"] >= meio - 3]
        if direita:
            inicios_direita[round(min(direita) / 4)] += 1

    # Duas colunas = (1) poucas linhas atravessam o meio da página e
    # (2) várias linhas do lado direito começam alinhadas na mesma margem.
    poucas_cruzam = linhas_que_cruzam <= len(linhas_visuais) * 0.2
    margem_direita = inicios_direita.most_common(1)[0][1] if inicios_direita else 0
    if not (poucas_cruzam and margem_direita >= 5):
        return [pagina.bbox]

    cruzam = [p for p in palavras if cruza_meio(p)]

    # Faixas horizontais ocupadas por textos que atravessam o meio.
    faixas = []
    for p in sorted(cruzam, key=lambda p: p["top"]):
        a, b = p["top"] - 1, p["bottom"] + 1
        if faixas and a <= faixas[-1][1]:
            faixas[-1][1] = max(faixas[-1][1], b)
        else:
            faixas.append([a, b])

    regioes = []
    cursor = topo
    for a, b in faixas + [[base, base]]:
        a, b = max(a, topo), min(b, base)
        if a - cursor > 2:   # trecho em duas colunas antes desta faixa
            regioes.append((x0, cursor, meio, a))
            regioes.append((meio, cursor, x1, a))
        if b - a > 0:        # a própria faixa, lida de ponta a ponta
            regioes.append((x0, a, x1, b))
        cursor = max(cursor, b)
    return regioes or [pagina.bbox]


def regiao_do_ponto(regioes, x, y):
    """Diz em qual região (índice) fica o ponto (x, y)."""
    for i, (rx0, rtop, rx1, rbottom) in enumerate(regioes):
        if rx0 <= x <= rx1 and rtop <= y <= rbottom:
            return i
    # Se não caiu em nenhuma, usa a região mais próxima verticalmente.
    return min(range(len(regioes)), key=lambda i: abs(regioes[i][1] - y))


def ler_pdf(caminho, hash_arquivo):
    """Lê o PDF inteiro e devolve:
       - 'itens': lista, na ordem de leitura, de linhas de texto e de
         marcadores de imagem;
       - 'texto_inicial': texto bruto das 2 primeiras páginas (cabeçalho)."""
    itens = []
    texto_inicial = ""
    contagem_linhas = Counter()      # para descobrir cabeçalhos/rodapés repetidos
    paginas_linhas = []

    with pdfplumber.open(caminho) as pdf:
        total_paginas = len(pdf.pages)

        # Imagens repetidas em várias páginas costumam ser LOGOTIPOS.
        # Contamos quantas vezes cada imagem aparece para ignorá-las.
        repeticoes_imagem = Counter()
        for pagina in pdf.pages:
            for img in pagina.images:
                repeticoes_imagem[(img.get("name"), tuple(img.get("srcsize") or ()))] += 1

        for num_pag, pagina in enumerate(pdf.pages, start=1):
            if num_pag <= 2:
                texto_inicial += (pagina.extract_text() or "") + "\n"

            regioes = dividir_em_regioes(pagina)
            # Cada elemento: (região, posição vertical, tipo, conteúdo)
            elementos = []

            # --- Linhas de texto, região por região ---
            for num_reg, area in enumerate(regioes):
                recorte = pagina.crop(area)
                for linha in recorte.extract_text_lines(return_chars=False):
                    texto = linha["text"].strip()
                    if texto:
                        # Números sozinhos no topo ou no rodapé da página são
                        # números de PÁGINA, não de questão -> descartamos.
                        na_margem = linha["top"] < pagina.height * 0.06 or linha["top"] > pagina.height * 0.92
                        if na_margem and re.fullmatch(r"\d{1,3}", texto):
                            continue
                        elementos.append((num_reg, linha["top"], "texto", texto))

            # --- Imagens: cada uma vai para a região onde começa ---
            if SALVAR_IMAGENS:
                for img in pagina.images:
                    chave_img = (img.get("name"), tuple(img.get("srcsize") or ()))
                    if total_paginas > 2 and repeticoes_imagem[chave_img] >= 3:
                        continue   # provável logotipo -> ignora
                    caminho_img = salvar_imagem(pagina, img, hash_arquivo, num_pag)
                    if caminho_img:
                        meio_x = (img["x0"] + img["x1"]) / 2
                        num_reg = regiao_do_ponto(regioes, meio_x, img["top"] + 1)
                        elementos.append((num_reg, img["top"], "imagem", caminho_img))

            # Ordena: primeiro pela região, depois de cima para baixo.
            elementos.sort(key=lambda e: (e[0], e[1]))
            itens_pagina = [{"tipo": t, "conteudo": c, "pagina": num_pag} for _, _, t, c in elementos]

            linhas_texto = {i["conteudo"] for i in itens_pagina if i["tipo"] == "texto"}
            contagem_linhas.update(linhas_texto)
            paginas_linhas.append(itens_pagina)

    # Remove cabeçalhos/rodapés: linhas idênticas presentes em muitas páginas.
    limite = max(3, int(len(paginas_linhas) * 0.4))
    # (Linhas que são só um número, como "3", não entram nessa regra:
    #  poderiam ser números de questão.)
    repetidas = {l for l, n in contagem_linhas.items()
                 if n >= limite and not l.isdigit()} if len(paginas_linhas) >= 3 else set()

    for itens_pagina in paginas_linhas:
        for item in itens_pagina:
            if item["tipo"] == "texto":
                t = item["conteudo"]
                if t in repetidas or eh_lixo(t):
                    continue
            itens.append(item)

    return itens, texto_inicial


def salvar_imagem(pagina, img, hash_arquivo, num_pag):
    """Recorta a área da imagem na página, reduz o tamanho e salva
    como JPG leve na pasta 'imagens'. Devolve o caminho do arquivo."""
    largura = img["x1"] - img["x0"]
    altura = img["bottom"] - img["top"]
    if largura < IMAGEM_TAMANHO_MINIMO or altura < IMAGEM_TAMANHO_MINIMO:
        return None   # muito pequena (ícone, marcador) -> ignora

    # Garante que o recorte fique dentro dos limites da página.
    px0, ptop, px1, pbottom = pagina.bbox
    area = (max(img["x0"], px0), max(img["top"], ptop),
            min(img["x1"], px1), min(img["bottom"], pbottom))
    if area[2] - area[0] < 5 or area[3] - area[1] < 5:
        return None

    try:
        imagem = pagina.crop(area).to_image(resolution=IMAGEM_RESOLUCAO).original
        imagem = imagem.convert("RGB")
        if imagem.width > IMAGEM_LARGURA_MAXIMA:
            proporcao = IMAGEM_LARGURA_MAXIMA / imagem.width
            imagem = imagem.resize((IMAGEM_LARGURA_MAXIMA, int(imagem.height * proporcao)))
        PASTA_IMAGENS.mkdir(exist_ok=True)
        nome = f"{hash_arquivo[:12]}_p{num_pag}_{int(area[1])}_{int(area[0])}.jpg"
        destino = PASTA_IMAGENS / nome
        imagem.save(destino, "JPEG", quality=IMAGEM_QUALIDADE_JPG, optimize=True)
        return f"imagens/{nome}"
    except Exception as erro:   # se algo der errado, só avisa e continua
        print(f"      (aviso) não foi possível salvar uma imagem da página {num_pag}: {erro}")
        return None


# =====================================================================
#  IDENTIFICAÇÃO DO CABEÇALHO (Banca, Ano, Órgão, Cargo)
# =====================================================================

def identificar_cabecalho(texto_inicial, nome_arquivo):
    """Procura Banca, Ano, Órgão e Cargo nas primeiras páginas e no
    nome do arquivo. O que não encontrar fica 'Não Identificado'."""
    texto_busca = sem_acento(texto_inicial + "\n" + Path(nome_arquivo).stem.replace("_", " "))
    linhas = [l.strip() for l in texto_inicial.splitlines() if l.strip()]

    # ---- BANCA: procura cada nome conhecido como palavra inteira ----
    banca = NAO_IDENTIFICADO
    melhor_posicao = None
    for nome_oficial, variacoes in BANCAS.items():
        for v in variacoes:
            achado = re.search(rf"\b{re.escape(v)}\b", texto_busca)
            if achado and (melhor_posicao is None or achado.start() < melhor_posicao):
                banca, melhor_posicao = nome_oficial, achado.start()

    # ---- ANO: o ano que mais aparece (priorizando o nome do arquivo) ----
    ano = NAO_IDENTIFICADO
    anos_arquivo = RE_ANO.findall(nome_arquivo)
    if anos_arquivo:
        ano = anos_arquivo[0]
    else:
        anos = RE_ANO.findall(texto_inicial)
        if anos:
            ano = Counter(anos).most_common(1)[0][0]

    # ---- ÓRGÃO: primeira linha curta contendo uma palavra típica ----
    orgao = NAO_IDENTIFICADO
    for linha in linhas[:60]:
        l_norm = sem_acento(linha)
        if len(linha) <= 120 and any(p in l_norm for p in PALAVRAS_ORGAO):
            orgao = re.sub(r"\s+", " ", linha).strip(" -–:")
            break

    # ---- CARGO: primeiro tenta "Cargo: ...", depois palavras típicas ----
    cargo = NAO_IDENTIFICADO
    for linha in linhas[:80]:
        m = RE_CARGO_ROTULO.search(linha)
        if m and len(m.group(1).strip()) > 2:
            cargo = m.group(1).strip(" -–:")
            break
    if cargo == NAO_IDENTIFICADO:
        for linha in linhas[:60]:
            l_norm = sem_acento(linha)
            if len(linha) <= 100 and any(re.match(rf"^\W*{p}\b", l_norm) for p in PALAVRAS_CARGO):
                cargo = linha.strip(" -–:")
                break

    return {"banca": banca, "ano": ano, "orgao": orgao, "cargo": cargo}


# =====================================================================
#  DISCIPLINA E ASSUNTO
# =====================================================================

def titulo_de_disciplina(linha):
    """Se a linha for um TÍTULO DE SEÇÃO (ex.: 'LÍNGUA PORTUGUESA' ou
    'Conhecimentos de Informática – Questões 21 a 30'), devolve o nome
    da disciplina. Caso contrário, devolve None."""
    l_norm = re.sub(r"[^A-Z ]+", " ", sem_acento(linha))
    l_norm = re.sub(r"\s+", " ", l_norm).strip()
    if not l_norm or len(l_norm.split()) > 9:
        return None
    # Títulos geralmente estão em MAIÚSCULAS ou são bem curtos
    letras = [c for c in linha if c.isalpha()]
    em_maiusculas = letras and sum(c.isupper() for c in letras) / len(letras) > 0.8
    if not em_maiusculas and len(l_norm.split()) > 5:
        return None
    l_norm = re.sub(r"^(CONHECIMENTOS|NOCOES) (BASICOS |GERAIS |ESPECIFICOS )?(DE |EM )?", "", l_norm)
    for nome_oficial, variacoes in DISCIPLINAS.items():
        for v in sorted(variacoes, key=len, reverse=True):
            if l_norm == v or l_norm.startswith(v + " "):
                if em_maiusculas or l_norm == v:
                    return nome_oficial
    return None


def adivinhar_disciplina(texto):
    """Quando a prova não tem títulos de seção, tenta adivinhar a
    disciplina pelas palavras-pista. Exige pelo menos 2 pistas."""
    t = sem_acento(texto).lower()
    pontos = {d: sum(t.count(p) for p in pistas) for d, pistas in PISTAS_DISCIPLINA.items()}
    melhor = max(pontos, key=pontos.get)
    return melhor if pontos[melhor] >= 2 else NAO_IDENTIFICADO


def adivinhar_assunto(disciplina, texto):
    """Procura palavras-chave de assunto dentro da disciplina."""
    if disciplina not in ASSUNTOS:
        return NAO_IDENTIFICADO
    t = sem_acento(texto).lower()
    pontos = {a: sum(t.count(p) for p in palavras) for a, palavras in ASSUNTOS[disciplina].items()}
    melhor = max(pontos, key=pontos.get)
    return melhor if pontos[melhor] >= 1 else NAO_IDENTIFICADO


# =====================================================================
#  GABARITO
# =====================================================================

def normalizar_resposta(resp):
    """Padroniza a resposta: 'certo' -> 'C', 'anulada' -> 'Anulada'..."""
    r = sem_acento(resp).strip()
    if r in ("ANULADA", "ANULADO", "X", "*"):
        return "Anulada"
    if r == "CERTO":
        return "C"
    if r == "ERRADO":
        return "E"
    return r


def ler_pares_gabarito(linhas):
    """Lê as respostas de um trecho de gabarito. Entende dois formatos:
       1) Pares na mesma linha:  '1-A  2-C  3-E'  ou  '01 B'
       2) Tabela em duas linhas: '1 2 3 4 5'  e logo abaixo  'A C B D E'"""
    respostas = {}
    for i, linha in enumerate(linhas):
        # Formato tabela: linha só de números seguida de linha só de letras
        numeros = linha.split()
        if numeros and all(n.isdigit() for n in numeros) and i + 1 < len(linhas):
            letras = linhas[i + 1].split()
            if len(letras) == len(numeros) and all(
                    re.fullmatch(r"[A-Ea-e]|X|\*|ANULADA|CERTO|ERRADO", l, re.I) for l in letras):
                for n, l in zip(numeros, letras):
                    respostas[int(n)] = normalizar_resposta(l)
                continue
        # Formato pares
        for n, l in RE_PAR_GABARITO.findall(linha):
            respostas.setdefault(int(n), normalizar_resposta(l))
    return respostas


def eh_arquivo_de_gabarito(nome_arquivo):
    return "GABARITO" in sem_acento(nome_arquivo)


# =====================================================================
#  SEPARAÇÃO DAS QUESTÕES (o "coração" do programa)
# =====================================================================

def separar_questoes(itens):
    """Percorre as linhas do PDF em ordem e monta as questões.

    Lógica (simplificada):
      - Uma linha que começa com 'Questão N' ou 'N.' / 'N)' inicia uma
        nova questão (desde que N seja o próximo número esperado, o que
        evita confundir com listas numeradas dentro do enunciado).
      - Linhas que começam com A), B), C)... (na ordem) são alternativas.
      - Outras linhas são continuação do enunciado ou da alternativa atual.
      - Títulos de seção (ex.: 'MATEMÁTICA') definem a disciplina das
        questões seguintes.
      - Ao encontrar a palavra 'GABARITO' no fim da prova, lê as respostas.
    """
    linhas_texto = [i["conteudo"] for i in itens if i["tipo"] == "texto"]

    # A prova é de múltipla escolha? (conta linhas com cara de alternativa)
    qtd_alt = sum(1 for l in linhas_texto if re.match(r"^\s*\(?[A-Ea-e]\s*[).]", l))
    multipla_escolha = qtd_alt >= 8

    # Procura um bloco de GABARITO no fim do documento.
    gabarito_interno = {}
    indice_gabarito = None
    for idx in range(len(itens) - 1, -1, -1):
        item = itens[idx]
        if item["tipo"] == "texto" and re.match(r"^\W*GABARITO", sem_acento(item["conteudo"])):
            trecho = [i["conteudo"] for i in itens[idx:] if i["tipo"] == "texto"]
            pares = ler_pares_gabarito(trecho)
            if len(pares) >= 3:
                gabarito_interno, indice_gabarito = pares, idx
            break
    if indice_gabarito is not None:
        itens = itens[:indice_gabarito]

    questoes = []
    atual = None                       # questão sendo montada no momento
    disciplina_atual = NAO_IDENTIFICADO
    ultimo_numero = 0

    def fechar_questao():
        if atual is not None:
            questoes.append(atual)

    for item in itens:
        # ---- Imagem: pertence à questão atual ----
        if item["tipo"] == "imagem":
            if atual is not None:
                atual["imagens"].append(item["conteudo"])
            continue

        linha = item["conteudo"]

        # ---- Título de disciplina? ----
        disc = titulo_de_disciplina(linha)
        if disc and (atual is None or atual["letras"] or not multipla_escolha):
            disciplina_atual = disc
            continue

        # ---- Início de nova questão? ----
        numero, resto = None, ""
        m = RE_QUESTAO_EXTENSO.match(linha)
        if m:
            numero, resto = int(m.group(1)), m.group(2)
        else:
            m = RE_QUESTAO_NUMERO.match(linha)
            if m:
                n = int(m.group(1))
                # Em prova de múltipla escolha, só aceitamos um número "solto"
                # depois que a questão anterior já tem alternativas (assim uma
                # lista "1. ... 2. ..." dentro do enunciado não vira questão).
                pronta = atual is None or not multipla_escolha or len(atual["letras"]) >= 2
                # O número precisa ser o próximo da sequência. Na múltipla
                # escolha toleramos pular 1 número (caso uma questão falhe).
                if multipla_escolha and atual is not None:
                    esperado = ultimo_numero < n <= ultimo_numero + 2
                else:
                    esperado = n == ultimo_numero + 1
                if esperado and pronta:
                    numero, resto = n, m.group(2)

        if numero is not None:
            fechar_questao()
            atual = {
                "numero": numero,
                "disciplina": disciplina_atual,
                "enunciado": [resto] if resto else [],
                "alternativas": {},   # letra -> lista de linhas
                "letras": [],         # ordem das letras encontradas
                "imagens": [],
            }
            ultimo_numero = numero
            continue

        if atual is None:
            continue   # ainda estamos nas instruções da capa

        # ---- Alternativa? (precisa ser a próxima letra: A, depois B...) ----
        m = RE_ALTERNATIVA.match(linha)
        if m:
            letra = m.group(1).upper()
            proxima = "ABCDE"[len(atual["letras"])] if len(atual["letras"]) < 5 else None
            if letra == proxima and (letra != "A" or atual["enunciado"]):
                texto_alt = m.group(2)
                atual["letras"].append(letra)
                atual["alternativas"][letra] = []
                # Trata alternativas grudadas: "(A) xx (B) yy (C) zz"
                partes = RE_ALTERNATIVA_INLINE.split(" " + texto_alt)
                atual["alternativas"][letra].append(partes[0])
                for k in range(1, len(partes) - 1, 2):
                    l2 = partes[k].upper()
                    if len(atual["letras"]) < 5 and l2 == "ABCDE"[len(atual["letras"])]:
                        atual["letras"].append(l2)
                        atual["alternativas"][l2] = [partes[k + 1]]
                    else:
                        atual["alternativas"][atual["letras"][-1]].append(f"({partes[k]}) {partes[k + 1]}")
                continue

        # ---- Continuação do texto ----
        if atual["letras"]:
            atual["alternativas"][atual["letras"][-1]].append(linha)
        else:
            atual["enunciado"].append(linha)

    fechar_questao()

    # Converte as listas de linhas em textos corridos.
    resultado = []
    for q in questoes:
        enunciado = juntar_linhas(q["enunciado"])
        alternativas = {l: juntar_linhas(q["alternativas"][l]) for l in q["letras"]}
        if not enunciado:
            continue
        resultado.append({
            "numero": q["numero"],
            "disciplina": q["disciplina"],
            "enunciado": enunciado,
            "alternativas": alternativas,
            "imagens": q["imagens"],
        })
    return resultado, gabarito_interno


# =====================================================================
#  MODO QCONCURSOS (PDFs exportados do site www.qconcursos.com)
# =====================================================================
# Os PDFs do QConcursos têm um formato fixo e muito informativo. Cada
# questão começa assim:
#
#   1 Q4305417 Estatística > Estatística descritiva , Tipos de variáveis
#   Ano: 2026 Banca: VUNESP Órgão: PREVCOM Prova: VUNESP - 2026 - PREVCOM - Analista
#   (enunciado...)
#   A primeira alternativa      <- a letra vem em NEGRITO
#   B segunda alternativa
#
# e no fim do arquivo há uma página "Respostas" (1: E 2: B 3: A ...).
# Aqui aproveitamos tudo isso: Disciplina e Assunto vêm prontos, e o
# código "Q4305417" vira o ID único (a mesma questão em dois PDFs
# diferentes é gravada uma só vez).

# O número da questão ("1 ") às vezes não aparece; só o código "Q3586110".
RE_QC_NUMERO = re.compile(r"^(?:(\d{1,4})\s+)?(Q\d{4,})\b\s*(.*)$")
RE_QC_CABECALHO = re.compile(
    r"Ano:\s*(\d{4})\s*Banca:\s*(.*?)\s*Órgão:\s*(.*?)\s*Provas?:\s*(.*?)[\s.…]*$", re.IGNORECASE)


def eh_pdf_qconcursos(caminho):
    """Diz se o PDF foi exportado do QConcursos (olha as 2 primeiras páginas)."""
    with pdfplumber.open(caminho) as pdf:
        if not pdf.pages:
            return False
        texto = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:2])
    return "qconcursos.com" in texto.lower() and bool(re.search(r"^(\d+\s+)?Q\d{4,}\b", texto, re.M))


def _linhas_qconcursos(caminho, hash_arquivo):
    """Lê o PDF e devolve a lista de linhas (com fonte, tamanho e
    posição) e de imagens, na ordem da página (de cima para baixo)."""
    itens = []
    with pdfplumber.open(caminho) as pdf:
        repeticoes = Counter()
        for pagina in pdf.pages:
            for img in pagina.images:
                repeticoes[(img.get("name"), tuple(img.get("srcsize") or ()))] += 1
        for num_pag, pagina in enumerate(pdf.pages, start=1):
            elementos = []
            for linha in pagina.extract_text_lines(return_chars=True):
                texto = linha["text"].strip()
                if not texto or "qconcursos.com" in texto.lower():
                    continue
                c0 = linha["chars"][0]
                elementos.append({
                    "tipo": "texto", "texto": texto, "pagina": num_pag,
                    "x0": linha["x0"], "top": linha["top"],
                    "tamanho": round(c0["size"], 1),
                    "negrito": "BOLD" in c0["fontname"].upper(),
                })
            if SALVAR_IMAGENS:
                for img in pagina.images:
                    chave = (img.get("name"), tuple(img.get("srcsize") or ()))
                    if len(pdf.pages) > 2 and repeticoes[chave] >= 3:
                        continue
                    caminho_img = salvar_imagem(pagina, img, hash_arquivo, num_pag)
                    if caminho_img:
                        elementos.append({"tipo": "imagem", "texto": caminho_img,
                                          "pagina": num_pag, "top": img["top"]})
            elementos.sort(key=lambda e: e["top"])
            itens.extend(elementos)
    return itens


def _eh_marcador_alternativa(item):
    """A letra da alternativa (A–E) aparece em negrito, na margem esquerda."""
    return (item["tipo"] == "texto" and item["negrito"] and item["x0"] < 50
            and re.match(r"^[A-E](\s|$)", item["texto"]) is not None
            and not item["texto"].startswith("Ano:"))


def _separar_cargo(prova, banca, ano, orgao):
    """'VUNESP - 2026 - PREVCOM - Analista de Tecnologia' -> 'Analista de Tecnologia'"""
    resto = prova
    m = re.search(rf"\b{ano}\b\s*-\s*", resto)
    if m:
        resto = resto[m.end():]
    if orgao and resto.startswith(orgao):
        resto = resto[len(orgao):]
    resto = resto.strip(" -–")
    return resto or NAO_IDENTIFICADO


def _classificacao(texto):
    """'Estatística > Descritiva , Tipos de variáveis' ->
       ('Estatística', 'Descritiva; Tipos de variáveis')"""
    texto = re.sub(r"\s+", " ", texto).strip(" ,")
    disciplinas, assuntos = [], []
    for parte in re.split(r"\s+,\s+", texto):
        parte = parte.strip(" ,")
        if not parte:
            continue
        if ">" in parte:
            disc, _, ass = parte.partition(">")
            disc, ass = disc.strip(), ass.strip()
            if disc and disc not in disciplinas:
                disciplinas.append(disc)
            if ass:
                assuntos.append(ass)
        else:
            assuntos.append(parte)
    return ("; ".join(disciplinas) or NAO_IDENTIFICADO,
            "; ".join(dict.fromkeys(assuntos)) or NAO_IDENTIFICADO)


def separar_questoes_qconcursos(itens):
    """Monta as questões a partir das linhas de um PDF do QConcursos."""
    # 1) Encontra o bloco "Respostas" (gabarito) no fim do arquivo.
    gabarito = {}
    for i, it in enumerate(itens):
        if it["tipo"] == "texto" and it["texto"].strip().lower() == "respostas":
            trecho = [x["texto"] for x in itens[i + 1:] if x["tipo"] == "texto"]
            gabarito = ler_pares_gabarito(trecho)
            itens = itens[:i]
            break

    # 2) Localiza as linhas de início de questão ("1 Q4305417 ...").
    inicios = [i for i, it in enumerate(itens)
               if it["tipo"] == "texto" and it["negrito"] and it["x0"] < 60
               and RE_QC_NUMERO.match(it["texto"])]

    # 3) A classificação (Disciplina > Assunto) às vezes quebra em linhas
    #    logo ACIMA e ABAIXO da linha do número. São linhas pequenas,
    #    à direita (x > 100), sem negrito, próximas da linha do número.
    #    Cada uma dessas linhas vai para a linha de número mais próxima.
    dono_classificacao = {}
    for j, it in enumerate(itens):
        if it["tipo"] != "texto" or it["negrito"] or it["x0"] <= 100 or it["tamanho"] >= 9:
            continue
        candidatos = [i for i in inicios if itens[i]["pagina"] == it["pagina"]
                      and abs(itens[i]["top"] - it["top"]) <= 45]
        if candidatos:
            dono_classificacao[j] = min(candidatos, key=lambda i: abs(itens[i]["top"] - it["top"]))

    questoes = []
    for k, i in enumerate(inicios):
        fim = inicios[k + 1] if k + 1 < len(inicios) else len(itens)
        m = RE_QC_NUMERO.match(itens[i]["texto"])
        numero = int(m.group(1)) if m.group(1) else None
        qid = m.group(2)

        # Classificação: junta as partes na ordem vertical.
        partes = [(itens[i]["top"], m.group(3))]
        partes += [(itens[j]["top"], itens[j]["texto"]) for j, d in dono_classificacao.items() if d == i]
        # (cada quebra de linha também separa assuntos)
        disciplina, assunto = _classificacao(" , ".join(t for _, t in sorted(partes) if t.strip()))

        # Corpo da questão (ignorando as linhas de classificação).
        corpo = [itens[j] for j in range(i + 1, fim) if j not in dono_classificacao]

        # Classificações muito longas podem ter linhas extras antes do "Ano:".
        for pos_ano in range(min(5, len(corpo))):
            if corpo[pos_ano]["tipo"] == "texto" and corpo[pos_ano]["texto"].startswith("Ano:"):
                extras = [c["texto"] for c in corpo[:pos_ano] if c["tipo"] == "texto" and c["tamanho"] < 9]
                if len(extras) == pos_ano:
                    disciplina, assunto = _classificacao(" , ".join(
                        [t for _, t in sorted(partes) if t.strip()] + extras))
                    corpo = corpo[pos_ano:]
                break

        # Cabeçalho "Ano: ... Banca: ... Órgão: ... Prova: ..." (pode quebrar linha)
        cabecalho, pos = "", 0
        while pos < len(corpo) and corpo[pos]["tipo"] == "texto" and (
                corpo[pos]["texto"].startswith("Ano:") or (cabecalho and corpo[pos]["tamanho"] < 9)):
            cabecalho += " " + corpo[pos]["texto"]
            pos += 1
        corpo = corpo[pos:]
        cab = RE_QC_CABECALHO.search(cabecalho.strip())
        if cab:
            ano, banca, orgao, prova = (g.strip() for g in cab.groups())
            cargo = _separar_cargo(prova, banca, ano, orgao)
        else:
            ano = banca = orgao = cargo = NAO_IDENTIFICADO
            prova = ""

        # Separa enunciado e alternativas usando a posição das letras.
        marcadores = [it for it in corpo if _eh_marcador_alternativa(it)]
        alternativas = {}            # letra -> lista de (página, posição, texto)
        enunciado, imagens = [], []
        letra_atual = None
        for it in corpo:
            if it["tipo"] == "imagem":
                imagens.append(it["texto"])
                continue
            if it in marcadores:
                letra_atual = it["texto"][0]
                alternativas.setdefault(letra_atual, [])
                resto = it["texto"][1:].strip()
                if resto:
                    alternativas[letra_atual].append((it["pagina"], it["top"], resto))
                continue
            # Itens de Certo/Errado (Cebraspe)
            if it["texto"] in ("Certo", "Errado") and it["x0"] >= 50:
                alternativas[it["texto"][0]] = [(it["pagina"], it["top"], it["texto"])]
                letra_atual = it["texto"][0]
                continue
            # Linha de alternativa longa: a letra fica sozinha, "no meio"
            # das linhas. Procuramos a letra mais próxima verticalmente.
            perto = [mk for mk in marcadores if mk["pagina"] == it["pagina"]
                     and len(mk["texto"].strip()) == 1 and abs(mk["top"] - it["top"]) <= 10]
            if perto and it["x0"] >= 50:
                letra = min(perto, key=lambda mk: abs(mk["top"] - it["top"]))["texto"][0]
                alternativas.setdefault(letra, []).append((it["pagina"], it["top"], it["texto"]))
            elif letra_atual:
                alternativas[letra_atual].append((it["pagina"], it["top"], it["texto"]))
            else:
                enunciado.append(it["texto"])

        alts = {}
        for letra in sorted(alternativas):
            linhas = [t for _, _, t in sorted(alternativas[letra])]
            alts[letra] = juntar_linhas(linhas)

        questoes.append({
            "id": qid, "numero": numero, "banca": banca, "ano": ano, "orgao": orgao,
            "cargo": cargo, "prova": prova, "disciplina": disciplina, "assunto": assunto,
            "enunciado": juntar_linhas(enunciado), "alternativas": alts,
            "imagens": imagens,
        })

    # Se o PDF não trouxe os números das questões, numeramos pela ordem,
    # alinhando com os números da página "Respostas" (ex.: 381, 382...).
    if any(q["numero"] is None for q in questoes):
        numeros_resp = sorted(gabarito)
        if len(numeros_resp) == len(questoes):
            for q, n in zip(questoes, numeros_resp):
                q["numero"] = n
        else:
            for pos, q in enumerate(questoes, start=1):
                q["numero"] = pos
    for q in questoes:
        q["gabarito"] = gabarito.get(q["numero"], NAO_IDENTIFICADO)
    return questoes


def processar_pdf_qconcursos(con, caminho, hash_arquivo):
    """Grava no banco as questões de um PDF do QConcursos."""
    nome = caminho.name
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    questoes = separar_questoes_qconcursos(_linhas_qconcursos(caminho, hash_arquivo))
    novas = 0
    for q in questoes:
        alternativas = json.dumps(q["alternativas"], ensure_ascii=False)
        cur = con.execute(
            """INSERT OR IGNORE INTO questoes
               (id, hash_conteudo, banca, ano, orgao, cargo, disciplina, assunto,
                numero, enunciado, alternativas, gabarito, imagens,
                arquivo_origem, chave_prova, data_inclusao)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (q["id"], hash_de_texto(q["id"]), q["banca"], q["ano"], q["orgao"], q["cargo"],
             q["disciplina"], q["assunto"], q["numero"], q["enunciado"], alternativas,
             q["gabarito"], json.dumps(q["imagens"], ensure_ascii=False), nome, q["prova"], agora),
        )
        if cur.rowcount:
            novas += 1
        elif q["gabarito"] != NAO_IDENTIFICADO:
            # Questão já existia: só completa o gabarito, se estiver faltando.
            con.execute("UPDATE questoes SET gabarito=? WHERE id=? AND gabarito=?",
                        (q["gabarito"], q["id"], NAO_IDENTIFICADO))
    con.execute("INSERT OR REPLACE INTO arquivos_processados VALUES (?,?,?,?,?,?)",
                (hash_arquivo, nome, "prova_qconcursos", chave_da_prova(nome), novas, agora))
    com_gab = sum(1 for q in questoes if q["gabarito"] != NAO_IDENTIFICADO)
    repetidas = len(questoes) - novas
    print(f"      [QConcursos] {len(questoes)} questões ({com_gab} com gabarito), "
          f"{novas} novas" + (f", {repetidas} já existiam no banco" if repetidas else "") + ".")
    return novas


# =====================================================================
#  PROCESSAMENTO DE UM ARQUIVO
# =====================================================================

def processar_pdf(con, caminho, hash_arquivo):
    """Lê um PDF e grava suas questões (ou seu gabarito) no banco."""
    nome = caminho.name
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    chave = chave_da_prova(nome)

    # PDFs do QConcursos têm um formato próprio, lido de forma especial.
    if eh_pdf_qconcursos(caminho):
        return processar_pdf_qconcursos(con, caminho, hash_arquivo)

    itens, texto_inicial = ler_pdf(caminho, hash_arquivo)
    linhas_texto = [i["conteudo"] for i in itens if i["tipo"] == "texto"]

    # PDF escaneado (só imagem, sem texto) -> não há o que ler.
    if len("".join(linhas_texto)) < 15:
        print("      ATENÇÃO: este PDF parece ser escaneado (sem texto selecionável).")
        print("      Ele precisaria passar por OCR antes. Veja o LEIAME.md.")
        con.execute("INSERT OR REPLACE INTO arquivos_processados VALUES (?,?,?,?,?,?)",
                    (hash_arquivo, nome, "sem_texto", chave, 0, agora))
        return 0

    # ---------- Caso 1: PDF de GABARITO ----------
    if eh_arquivo_de_gabarito(nome):
        respostas = ler_pares_gabarito(linhas_texto)
        definitivo = 1 if "DEFINITIVO" in sem_acento(nome) else 0
        for numero, resp in respostas.items():
            if definitivo:
                con.execute("INSERT OR REPLACE INTO gabaritos VALUES (?,?,?,?)",
                            (chave, numero, resp, 1))
            else:
                con.execute("INSERT OR IGNORE INTO gabaritos VALUES (?,?,?,?)",
                            (chave, numero, resp, 0))
        con.execute("INSERT OR REPLACE INTO arquivos_processados VALUES (?,?,?,?,?,?)",
                    (hash_arquivo, nome, "gabarito", chave, len(respostas), agora))
        print(f"      Gabarito com {len(respostas)} respostas (chave: '{chave}').")
        return 0

    # ---------- Caso 2: PDF de PROVA ----------
    cab = identificar_cabecalho(texto_inicial, nome)
    print(f"      Banca: {cab['banca']} | Ano: {cab['ano']} | Órgão: {cab['orgao']}")
    print(f"      Cargo: {cab['cargo']}")

    questoes, gabarito_interno = separar_questoes(itens)
    novas = 0
    for q in questoes:
        disciplina = q["disciplina"]
        texto_completo = q["enunciado"] + " " + " ".join(q["alternativas"].values())
        if disciplina == NAO_IDENTIFICADO:
            disciplina = adivinhar_disciplina(texto_completo)
        assunto = adivinhar_assunto(disciplina, texto_completo)

        alternativas = q["alternativas"]
        # Provas do tipo Certo/Errado (comum no Cebraspe) não têm A-E.
        if not alternativas and cab["banca"] == "Cebraspe":
            alternativas = {"C": "Certo", "E": "Errado"}

        gabarito = gabarito_interno.get(q["numero"], NAO_IDENTIFICADO)
        conteudo = q["enunciado"] + json.dumps(alternativas, ensure_ascii=False)
        h = hash_de_texto(conteudo)
        id_unico = f"Q-{hash_arquivo[:8]}-{q['numero']:03d}"

        cursor = con.execute(
            """INSERT OR IGNORE INTO questoes
               (id, hash_conteudo, banca, ano, orgao, cargo, disciplina, assunto,
                numero, enunciado, alternativas, gabarito, imagens,
                arquivo_origem, chave_prova, data_inclusao)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (id_unico, h, cab["banca"], cab["ano"], cab["orgao"], cab["cargo"],
             disciplina, assunto, q["numero"], q["enunciado"],
             json.dumps(alternativas, ensure_ascii=False), gabarito,
             json.dumps(q["imagens"], ensure_ascii=False), nome, chave, agora),
        )
        novas += cursor.rowcount

    con.execute("INSERT OR REPLACE INTO arquivos_processados VALUES (?,?,?,?,?,?)",
                (hash_arquivo, nome, "prova", chave, novas, agora))
    repetidas = len(questoes) - novas
    print(f"      {len(questoes)} questões encontradas, {novas} novas gravadas"
          + (f" ({repetidas} já existiam no banco)" if repetidas else "") + ".")
    if gabarito_interno:
        print(f"      Gabarito encontrado no próprio PDF ({len(gabarito_interno)} respostas).")
    return novas


def aplicar_gabaritos_separados(con):
    """Liga os PDFs de gabarito às provas pelo nome do arquivo.
    Ex.: 'TRT_2023_Analista_GABARITO.pdf' preenche as respostas das
    questões de 'TRT_2023_Analista.pdf'. Funciona em qualquer ordem."""
    total = 0
    for chave, numero, resposta, definitivo in con.execute(
            "SELECT chave_prova, numero, resposta, definitivo FROM gabaritos").fetchall():
        if definitivo:   # gabarito definitivo sempre prevalece
            cur = con.execute("UPDATE questoes SET gabarito=? WHERE chave_prova=? AND numero=? AND gabarito<>?",
                              (resposta, chave, numero, resposta))
        else:            # preliminar só preenche o que está vazio
            cur = con.execute("UPDATE questoes SET gabarito=? WHERE chave_prova=? AND numero=? AND gabarito=?",
                              (resposta, chave, numero, NAO_IDENTIFICADO))
        total += cur.rowcount
    if total:
        print(f"\n  Gabaritos separados aplicados a {total} questões.")


# =====================================================================
#  EXPORTAÇÃO (JSON e CSV para consulta)
# =====================================================================

def exportar(con):
    """Gera 'banco_questoes.json' e 'banco_questoes.csv' com TODAS as
    questões do banco (o CSV abre direto no Excel)."""
    con.row_factory = sqlite3.Row
    linhas = con.execute("SELECT * FROM questoes ORDER BY arquivo_origem, numero").fetchall()
    con.row_factory = None

    dados = []
    for r in linhas:
        dados.append({
            "id": r["id"],
            "banca": r["banca"],
            "ano": r["ano"],
            "orgao": r["orgao"],
            "cargo": r["cargo"],
            "disciplina": r["disciplina"],
            "assunto": r["assunto"],
            "numero": r["numero"],
            "enunciado": r["enunciado"],
            "alternativas": json.loads(r["alternativas"] or "{}"),
            "gabarito": r["gabarito"],
            "imagens": json.loads(r["imagens"] or "[]"),
            "arquivo_origem": r["arquivo_origem"],
        })

    with open(ARQUIVO_JSON, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    # CSV com ";" e codificação que o Excel brasileiro entende.
    with open(ARQUIVO_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["ID", "Banca", "Ano", "Órgão", "Cargo", "Disciplina", "Assunto", "Nº",
                    "Enunciado", "A", "B", "C", "D", "E", "Gabarito", "Imagens", "Arquivo"])
        for d in dados:
            alt = d["alternativas"]
            w.writerow([d["id"], d["banca"], d["ano"], d["orgao"], d["cargo"], d["disciplina"],
                        d["assunto"], d["numero"], d["enunciado"],
                        alt.get("A", ""), alt.get("B", ""), alt.get("C", ""), alt.get("D", ""),
                        alt.get("E", ""), d["gabarito"], ", ".join(d["imagens"]), d["arquivo_origem"]])
    return len(dados)


# =====================================================================
#  PROGRAMA PRINCIPAL
# =====================================================================

def main():
    global PASTA_PDFS
    print("=" * 65)
    print("  EXTRATOR DE QUESTÕES DE CONCURSOS")
    print("=" * 65)

    # Opcional: você pode informar outra pasta de PDFs ao rodar, ex.:
    #   python3 extrator_questoes.py "/caminho/da/minha/pasta"
    # (o banco, as imagens e as planilhas continuam ao lado deste programa)
    if len(sys.argv) > 1:
        PASTA_PDFS = Path(sys.argv[1]).expanduser()
        if not PASTA_PDFS.is_dir():
            print(f"\n[ERRO] Pasta não encontrada:\n  {PASTA_PDFS}")
            print("Dica: arraste a pasta para a janela do Terminal para colar o caminho certo.")
            return
    print(f"\nPasta dos PDFs: {PASTA_PDFS}")

    # iCloud: arquivos que estão só na nuvem aparecem como ".nome.pdf.icloud"
    # e não podem ser lidos até serem baixados para o computador.
    na_nuvem = list(PASTA_PDFS.rglob("*.pdf.icloud")) if PASTA_PDFS.exists() else []
    if na_nuvem:
        print(f"\nATENÇÃO: {len(na_nuvem)} PDF(s) estão só no iCloud e serão pulados.")
        print("No Finder, clique com o botão direito na pasta > 'Baixar Agora' e rode de novo.")

    # Cria a pasta de PDFs se ela ainda não existir.
    if not PASTA_PDFS.exists():
        PASTA_PDFS.mkdir()
        print(f"\nA pasta de PDFs foi criada em:\n  {PASTA_PDFS}")
        print("Coloque seus PDFs lá dentro e rode o programa de novo.")
        return

    # Lista os PDFs (inclusive dentro de subpastas). Gabaritos por último,
    # só por organização (a ordem não altera o resultado).
    pdfs = sorted((p for p in PASTA_PDFS.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf"),
                  key=lambda p: (eh_arquivo_de_gabarito(p.name), p.name.lower()))
    if not pdfs:
        print(f"\nNenhum PDF encontrado em:\n  {PASTA_PDFS}")
        return

    con = abrir_banco()
    ja_processados = {h for (h,) in con.execute("SELECT hash_arquivo FROM arquivos_processados")}

    novos, pulados, total_questoes = 0, 0, 0
    for caminho in pdfs:
        h = hash_de_arquivo(caminho)
        if h in ja_processados:
            pulados += 1
            continue
        novos += 1
        print(f"\n>> Processando: {caminho.name}")
        try:
            total_questoes += processar_pdf(con, caminho, h)
            con.commit()   # grava no disco a cada arquivo (seguro contra quedas)
            ja_processados.add(h)
        except Exception as erro:
            con.rollback()
            print(f"      ERRO ao ler este arquivo: {erro}")
            print("      O arquivo será tentado de novo na próxima execução.")

    aplicar_gabaritos_separados(con)
    con.commit()
    total_banco = exportar(con)
    con.close()

    print("\n" + "-" * 65)
    print(f"  PDFs novos processados .......: {novos}")
    print(f"  PDFs já lidos (pulados) ......: {pulados}")
    print(f"  Questões novas gravadas ......: {total_questoes}")
    print(f"  Total de questões no banco ...: {total_banco}")
    print("-" * 65)
    print(f"  Banco de dados: {ARQUIVO_BANCO.name}")
    print(f"  Planilha (Excel): {ARQUIVO_CSV.name}")
    print(f"  JSON: {ARQUIVO_JSON.name}")
    print("=" * 65)


# Esta linha faz o programa começar pela função main() quando você
# roda "python extrator_questoes.py".
if __name__ == "__main__":
    main()
