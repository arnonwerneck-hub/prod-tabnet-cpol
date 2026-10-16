# Extrator de Questões de Concursos — Guia passo a passo

Este programa lê provas de concurso em PDF e monta um banco de questões
com: **ID, Banca, Ano, Órgão, Cargo, Disciplina, Assunto, Enunciado,
Alternativas (A–E) e Gabarito**. Imagens das questões são recortadas e
salvas como JPG pequeno.

---

## 1. Instalar o Python (uma única vez)

### Windows
1. Acesse <https://www.python.org/downloads/> e clique em **Download Python 3.x**.
2. Abra o instalador. **IMPORTANTE:** na primeira tela, marque a caixa
   **"Add Python to PATH"** (ou "Add python.exe to PATH").
3. Clique em **Install Now** e aguarde terminar.
4. Para conferir: aperte a tecla Windows, digite `cmd`, abra o
   **Prompt de Comando** e digite:
   ```
   python --version
   ```
   Deve aparecer algo como `Python 3.12.x`.

### Mac
1. Baixe em <https://www.python.org/downloads/> e instale normalmente.
2. Abra o aplicativo **Terminal** e confira com `python3 --version`.
   (No Mac, use sempre `python3` e `pip3` no lugar de `python` e `pip`.)

---

## 2. Instalar a biblioteca necessária (uma única vez)

No Prompt de Comando (Windows) ou Terminal (Mac), digite:

```
pip install pdfplumber
```

(No Mac: `pip3 install pdfplumber`.)

O `pdfplumber` lê o texto dos PDFs com a posição de cada linha (isso
permite entender provas em duas colunas) e também recorta as imagens.
Ele instala automaticamente o que precisa (Pillow e pypdfium2).

---

## 3. Organizar as pastas

Coloque a pasta `extrator_questoes` onde preferir (ex.: Documentos).
Ela deve ficar assim:

```
extrator_questoes/
├── extrator_questoes.py    <- o programa
├── LEIAME.md               <- este guia
├── requirements.txt
└── pdfs/                   <- COLOQUE SEUS PDFs AQUI DENTRO
```

**Dicas para os nomes dos arquivos (ajudam muito):**
- Inclua o ano no nome: `TJRJ_2024_Analista.pdf`. O ano do nome do
  arquivo tem prioridade sobre o que está no texto.
- Para um gabarito em PDF separado, use **o mesmo nome da prova +
  a palavra "gabarito"**:
  - Prova: `TJRJ_2024_Analista.pdf`
  - Gabarito: `TJRJ_2024_Analista_gabarito.pdf`
    (ou `..._gabarito_definitivo.pdf` — o **definitivo** sempre
    substitui o preliminar).
- Pode criar subpastas dentro de `pdfs/`; o programa procura em todas.

---

## 4. Rodar o programa

1. Abra o Prompt de Comando / Terminal **dentro da pasta**
   `extrator_questoes`:
   - **Windows:** abra a pasta no Explorador de Arquivos, clique na
     barra de endereço, digite `cmd` e aperte Enter.
   - **Mac:** clique com o botão direito na pasta > **Novo Terminal na Pasta**.
2. Digite:
   ```
   python extrator_questoes.py
   ```
   (Mac: `python3 extrator_questoes.py`)

Você verá algo assim:

```
>> Processando: TJRJ_2024_Analista.pdf
      Banca: FGV | Ano: 2024 | Órgão: TRIBUNAL DE JUSTIÇA DO ESTADO DO RIO DE JANEIRO
      Cargo: Analista Judiciário - Área Judiciária
      60 questões encontradas, 60 novas gravadas.
      Gabarito encontrado no próprio PDF (60 respostas).
```

### Usar uma pasta que já existe (ex.: iCloud), sem copiar os PDFs
Informe o caminho da pasta depois do nome do programa, **entre aspas**:
```
python3 extrator_questoes.py "/Users/SEU_USUARIO/Library/Mobile Documents/com~apple~CloudDocs/QUESTÕES DE CONCURSOS"
```
Dica: digite `python3 extrator_questoes.py ` (com espaço no fim) e
**arraste a pasta do Finder para a janela do Terminal**: o caminho é
colado automaticamente. No iCloud, antes de rodar, clique com o botão
direito na pasta > **Baixar Agora**, para que os PDFs estejam no Mac.

### Adicionar novas provas depois
Basta colocar os novos PDFs na pasta `pdfs/` e rodar de novo o mesmo
comando. **Só os arquivos novos são lidos**; os antigos são pulados
(mesmo se você renomeá-los). Questões repetidas não são duplicadas.

---

## 5. O que o programa gera

| Arquivo | Para que serve |
|---|---|
| `banco_questoes.db` | **O banco principal** (SQLite). Não apague: é ele que "lembra" o que já foi lido. |
| `banco_questoes.csv` | Planilha com todas as questões. Abre direto no **Excel**. |
| `banco_questoes.json` | As mesmas questões em JSON (útil para outros programas/sites). |
| `imagens/` | Imagens recortadas das questões, em JPG leve. O campo "Imagens" diz qual é de qual questão. |

O CSV e o JSON são **recriados a cada execução** a partir do banco.

### Exemplo de uma questão no JSON
```json
{
  "id": "Q-737c83d2-001",
  "banca": "FGV",
  "ano": "2024",
  "orgao": "TRIBUNAL DE JUSTIÇA DO ESTADO DO RIO DE JANEIRO",
  "cargo": "Analista Judiciário - Área Judiciária",
  "disciplina": "Língua Portuguesa",
  "assunto": "Crase",
  "numero": 1,
  "enunciado": "Assinale a alternativa em que o uso da crase está correto...",
  "alternativas": {"A": "Fui à pé.", "B": "Refiro-me à você.", "C": "Cheguei à cidade.", "D": "Começou à chover.", "E": "Vou à Brasília."},
  "gabarito": "C",
  "imagens": [],
  "arquivo_origem": "TJRJ_2024_Analista.pdf"
}
```

---

## 6. Categorizar manualmente o que ficou "Não Identificado"

Instale o programa gratuito **DB Browser for SQLite**
(<https://sqlitebrowser.org/dl/>):

1. Abra o DB Browser > **Abrir banco de dados** > escolha `banco_questoes.db`.
2. Vá na aba **Navegar dados** e escolha a tabela `questoes`.
3. No filtro da coluna `disciplina` digite `Não Identificado`.
4. Clique na célula e digite o valor correto.
5. Clique em **Escrever modificações** (Ctrl+S) para salvar.

Na próxima execução do programa, o CSV/JSON já sairão com suas correções.
(Editar o CSV no Excel **não** altera o banco; edite pelo DB Browser.)

---

## 7. Personalizar o reconhecimento

No começo do arquivo `extrator_questoes.py`, na seção **CONFIGURAÇÕES**,
há listas que você pode editar com o Bloco de Notas:

- `BANCAS` — nomes de bancas e como aparecem no PDF;
- `PALAVRAS_ORGAO` e `PALAVRAS_CARGO` — palavras que indicam órgão/cargo;
- `DISCIPLINAS` — títulos de seção que definem a matéria;
- `PISTAS_DISCIPLINA` e `ASSUNTOS` — palavras-chave para adivinhar
  matéria e assunto quando a prova não tem títulos;
- `SALVAR_IMAGENS`, `IMAGEM_QUALIDADE_JPG`, `IMAGEM_LARGURA_MAXIMA` —
  controle do tamanho das imagens.

Depois de mudar as listas, para **reprocessar tudo** do zero: apague
`banco_questoes.db` e a pasta `imagens/`, e rode o programa de novo.

---

## 8. Como a extração funciona (resumo)

- **Cabeçalho:** procura nas 2 primeiras páginas e no nome do arquivo:
  banca (lista conhecida), ano (19xx/20xx), órgão (linhas com
  "Tribunal", "Prefeitura", "Secretaria"...), cargo ("Cargo: ..." ou
  linhas começando com "Analista", "Técnico"...).
- **Duas colunas:** detecta o corredor central e lê a coluna esquerda
  antes da direita; títulos que atravessam a página são lidos inteiros.
- **Questões:** linhas que começam com `Questão 12`, `12.`, `12)`,
  `12 -` ou só `12`. O número precisa seguir a sequência, para não
  confundir listas numeradas do enunciado com novas questões.
- **Alternativas:** linhas começando com `A)`, `(A)`, `a)`, `A.`,
  `A -`, na ordem A→E. Alternativas na mesma linha (`(A) x (B) y`)
  também são separadas.
- **Certo/Errado (Cebraspe):** itens sem alternativas recebem
  `{"C": "Certo", "E": "Errado"}`.
- **Gabarito:** bloco "GABARITO" no fim da prova, ou PDF separado com
  "gabarito" no nome. Entende `1-A 2-C`, `01 B` e tabelas
  (`1 2 3` / `A C B`). Questões anuladas ficam como `Anulada`.
- **Cabeçalhos/rodapés repetidos** em várias páginas e números de
  página são descartados. Logotipos repetidos não viram imagem.

---

## 8.1 PDFs exportados do QConcursos

O programa reconhece sozinho os PDFs gerados pelo site **QConcursos**
(têm "www.qconcursos.com" no topo) e usa um leitor próprio para eles:

- **ID único** = o código da questão no site (ex.: `Q4305417`). A mesma
  questão em vários PDFs é gravada uma única vez.
- **Banca, Ano, Órgão e Cargo** vêm da linha `Ano: ... Banca: ... Órgão: ... Prova: ...`.
- **Disciplina e Assunto** vêm da classificação do site
  (`Estatística > Tipos de variáveis , ...`). Vários assuntos ficam
  separados por `;`.
- **Gabarito** vem da página "Respostas" no fim do PDF. Ao exportar no
  QConcursos, **marque a opção de incluir o gabarito**; senão o campo
  fica "Não Identificado".
- Fórmulas, tabelas e gráficos que estão como imagem no PDF são salvos
  na pasta `imagens/` e ligados à questão.

## 9. Limitações e problemas comuns

| Situação | O que fazer |
|---|---|
| `'python' não é reconhecido...` | Reinstale o Python marcando **Add Python to PATH**. |
| `A biblioteca 'pdfplumber' não está instalada` | Rode `pip install pdfplumber`. |
| "este PDF parece ser escaneado" | O PDF é só uma foto, sem texto. Passe-o por um OCR antes (ex.: abrir no Google Drive com Google Docs, ou "Reconhecer texto" do Adobe Acrobat) e salve como novo PDF. |
| Poucas questões encontradas | A prova pode usar um formato diferente. Para numeração que não começa em 1 (ex.: caderno que começa na questão 31), a prova precisa ter "Questão 31" escrito, ou a numeração 1, 2, 3... |
| Gabarito não aplicado | Confira se o nome do PDF de gabarito é igual ao da prova + "gabarito". Provas com vários "tipos" de caderno (Tipo 1, Tipo 2) podem ter gabaritos diferentes por tipo: deixe só o gabarito do tipo correspondente. |
| Disciplina/Assunto "Não Identificado" | Normal quando a prova não tem títulos de seção. Categorize no DB Browser (item 6) ou acrescente palavras-chave nas listas (item 7). |
| Erro em um arquivo | O programa avisa, pula esse arquivo e continua os outros. Ele tentará de novo na próxima execução. |

Nenhuma extração automática é 100% perfeita, já que cada banca formata
de um jeito. Vale conferir uma amostra das questões de cada prova nova.
