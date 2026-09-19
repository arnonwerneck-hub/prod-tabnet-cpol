# Produção Ambulatorial — Policlínicas e Centros Cariocas (SMS-RJ)

Extração automatizada de dados do [TabNet Rio de Janeiro](https://tabnet.rio.rj.gov.br/cgi-bin/dh?sia/definicoes/producao_2008.def)
(SIA/SUS — Produção Ambulatorial de Procedimentos) para acompanhamento da produção de
Policlínicas e Centros Cariocas do Município do Rio de Janeiro em 2026 (Profissional-CBO
× Mês de Atendimento, Qtd. Apresentada e Qtd. Aprovada).

## Estrutura

```
scripts/
  extract_tabnet.py         # extrai os dados brutos do TabNet (Profissional-CBO x Mês, por unidade)
  consolidar_xlsx.py        # gera o .xlsx final e o CSV-resumo a partir do CSV bruto
  gerar_dados_dashboard.py  # gera o JSON (data.json.js) usado pelo dashboard
data/
  tabnet_raw_long.csv                          # dados brutos extraídos (formato longo)
  producao_policlinicas_2026_resumo_mensal.csv # CSV final organizado (unidade x mês, apresentada/aprovada/taxa)
  tabnet_checkpoint.txt                        # checkpoint da extração (gitignored)
dashboard/
  index.html                # dashboard interativo (Chart.js), publicado no Vercel
  data.json.js              # dados pré-agregados embutidos no dashboard
Producao_Policlinicas_RJ_2026.xlsx   # planilha final com todas as abas
```

## Unidades monitoradas

- SMS Policlínica José Paranhos Fontenelle — AP 31
- SMS Policlínica Hélio Pellegrino — AP 22
- SMS Policlínica Carlos Alberto Nascimento — AP 52
- SMS Policlínica Newton Alves Cardozo — AP 31
- SMS Policlínica Lincoln de Freitas Filho — AP 53
- SMS Policlínica Manoel Guilherme (PAM Bangu) — AP 51
- SMS Policlínica Newton Bethlem — AP 40
- SMS Policlínica Rocha Maia — AP 21
- SMS Policlínica Rodolpho Rocco — AP 32
- SMS Policlínica Antônio Ribeiro Netto — AP 10
- SMS Centro Carioca de Especialidade da Zona Oeste — AP 52
- SMS Centro Carioca de Reabilitação da Zona Oeste — AP 52
- SMS Centro Carioca de Hemodiálise — AP 52 (sem produção registrada por Profissional-CBO em 2026 até o momento — confirmado via a página "Nenhum registro selecionado" do próprio TabNet, não é falha de extração)

> A unidade "SMS Coord Geral do Super Centro Carioca de Saúde ZO AP 52" não foi localizada
> no cadastro de estabelecimentos do TabNet (busca por nome/variações não retornou
> correspondência) e não está incluída na extração.

## Período

Jan/2026 a Jul/2026 — é o período mais recente publicado pelo TabNet no momento da extração
(meses seguintes ainda não disponíveis na fonte). O script está preparado para estender o
intervalo assim que novos meses forem publicados (`TODOS_ARQUIVOS` em `extract_tabnet.py`).

## Como reprocessar os dados

```bash
python3 scripts/extract_tabnet.py        # poucos minutos; salva checkpoint incremental
python3 scripts/consolidar_xlsx.py        # gera o .xlsx e o CSV-resumo
python3 scripts/gerar_dados_dashboard.py  # gera dashboard/data.json.js
```

O script de extração salva checkpoint em `data/tabnet_checkpoint.txt` — pode ser
interrompido e retomado sem reprocessar tarefas já concluídas. Consultas de vários meses
que retornam uma resposta truncada/ambígua (reset de conexão do servidor do TabNet) são
divididas automaticamente em consultas mensais e só são aceitas como "zero" quando a
página confirma explicitamente "Nenhum registro selecionado" — isso evita perder dados
reais silenciosamente.

## Dashboard

`dashboard/index.html` é um dashboard estático (sem backend), com:

- Filtros por unidade, ano e mês (linha de controles no cabeçalho).
- KPIs de produtividade: totais e variação mês a mês, taxa de aprovação (meter),
  produtividade média mensal, melhor mês, unidade líder, unidades monitoradas e
  profissionais-CBO ativos.
- **Diferença mensal** — gráfico de barras divergente da glosa (apresentada − aprovada)
  por mês, para destacar variações que ficam escondidas em duas linhas quase sobrepostas.
- **Ranking por unidade — Taxa de aprovação**, em ordem decrescente, com cores por
  faixa (≥90% boa / 70–89% atenção / <70% crítica).
- **Ranking por unidade — Volume apresentado**, também decrescente.
- Top Profissional-CBO da rede (período completo).
- Mapa de calor Unidade × Mês (intensidade de produção).
- Tabela detalhada, ordenável, por unidade e mês.
- Modo claro/escuro (preferência do sistema + alternância manual).

Publicado no Vercel a partir da pasta `dashboard/`. Para testar localmente:

```bash
python3 -m http.server 8743 --directory dashboard
```

## Fonte dos dados

Secretaria Municipal de Saúde do Rio de Janeiro — Sistema de Informações Ambulatoriais
do SUS (SIA/SUS), via TabNet.
