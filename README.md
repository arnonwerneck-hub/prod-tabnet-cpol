# Produção Ambulatorial — Policlínicas e Centros Cariocas (SMS-RJ)

Extração automatizada de dados do [TabNet Rio de Janeiro](https://tabnet.rio.rj.gov.br/cgi-bin/dh?sia/definicoes/producao_2008.def)
(SIA/SUS — Produção Ambulatorial de Procedimentos) para acompanhamento da produção de
Policlínicas e Centros Cariocas do Município do Rio de Janeiro, a partir de 2023.

## Estrutura

```
scripts/
  extract_tabnet.py         # extrai os dados brutos do TabNet (Profissional-CBO x Mês, por unidade)
  consolidar_xlsx.py        # gera o arquivo .xlsx final a partir do CSV bruto
  gerar_dados_dashboard.py  # gera o JSON (data.json.js) usado pelo dashboard
data/
  tabnet_raw_long.csv       # dados brutos extraídos (formato longo)
dashboard/
  index.html                # dashboard interativo (Chart.js), publicado no Vercel
  data.json.js              # dados pré-agregados embutidos no dashboard
Producao_Policlinicas_RJ_2023_em_diante.xlsx   # planilha final com todas as abas
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
- SMS Centro Carioca de Hemodiálise — AP 52

> A unidade "SMS Coord Geral do Super Centro Carioca de Saúde ZO AP 52" não foi localizada
> no cadastro de estabelecimentos do TabNet (busca por nome/variações não retornou
> correspondência) e não está incluída na extração.

## Como reprocessar os dados

```bash
python3 scripts/extract_tabnet.py        # ~45-70 min (limitado pela velocidade do servidor do TabNet)
python3 scripts/consolidar_xlsx.py        # gera o .xlsx
python3 scripts/gerar_dados_dashboard.py  # gera dashboard/data.json.js
```

O script de extração salva checkpoint em `data/tabnet_checkpoint.txt` — pode ser
interrompido e retomado sem reprocessar tarefas já concluídas.

## Dashboard

`dashboard/index.html` é um dashboard estático (sem backend), com filtro por unidade,
evolução mensal (Apresentada x Aprovada), ranking por unidade, taxa de aprovação e
distribuição por Profissional-CBO. Publicado no Vercel a partir da pasta `dashboard/`.

## Fonte dos dados

Secretaria Municipal de Saúde do Rio de Janeiro — Sistema de Informações Ambulatoriais
do SUS (SIA/SUS), via TabNet.
