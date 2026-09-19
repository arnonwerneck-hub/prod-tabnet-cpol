#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Le data/tabnet_raw_long.csv e gera um JSON compacto e pre-agregado para
alimentar o dashboard estatico (public/data.json).
"""
import json
import os
import re

import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IN_CSV = os.path.join(BASE_DIR, "data", "tabnet_raw_long.csv")
OUT_JS = os.path.join(BASE_DIR, "dashboard", "data.json.js")

MES_ORDEM = {m: i for i, m in enumerate(["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"])}


def mes_chave(label):
    m = re.match(r"([A-Za-z]{3})/(\d{4})", label)
    if not m:
        return (9999, 99)
    mes, ano = m.group(1), int(m.group(2))
    return (ano, MES_ORDEM.get(mes, 99))


def abreviar_nome(nome):
    nome = nome.replace("SMS POLICLINICA ", "").replace("SMS CENTRO CARIOCA DE ", "C.C. ")
    return nome.strip()


def main():
    df = pd.read_csv(IN_CSV, encoding="utf-8")
    df = df[df["mes_atendimento"].apply(lambda x: mes_chave(x)[0] >= 2023)].copy()
    df["mes_ord"] = df["mes_atendimento"].apply(mes_chave)

    meses = sorted(df["mes_atendimento"].unique(), key=mes_chave)
    estabelecimentos = sorted(df["estabelecimento_nome"].unique())

    # 1) Serie mensal por estabelecimento e conteudo (apresentada/aprovada)
    monthly = (
        df.groupby(["estabelecimento_nome", "mes_atendimento", "conteudo"], as_index=False)["valor"]
        .sum()
    )
    pivot = monthly.pivot_table(
        index=["estabelecimento_nome", "mes_atendimento"], columns="conteudo", values="valor", fill_value=0
    ).reset_index()
    for c in ("apresentada", "aprovada"):
        if c not in pivot.columns:
            pivot[c] = 0
    pivot["mes_ord"] = pivot["mes_atendimento"].apply(mes_chave)
    pivot = pivot.sort_values(["estabelecimento_nome", "mes_ord"])

    monthly_records = []
    for _, r in pivot.iterrows():
        apresentada = float(r["apresentada"])
        aprovada = float(r["aprovada"])
        taxa = round(100.0 * aprovada / apresentada, 1) if apresentada else 0.0
        monthly_records.append(
            {
                "estabelecimento": r["estabelecimento_nome"],
                "mes": r["mes_atendimento"],
                "apresentada": apresentada,
                "aprovada": aprovada,
                "taxa_aprovacao": taxa,
            }
        )

    # 2) Totais por estabelecimento (para ranking / barras)
    totals = (
        df.groupby(["estabelecimento_nome", "conteudo"], as_index=False)["valor"].sum()
        .pivot_table(index="estabelecimento_nome", columns="conteudo", values="valor", fill_value=0)
        .reset_index()
    )
    for c in ("apresentada", "aprovada"):
        if c not in totals.columns:
            totals[c] = 0
    totals_records = []
    for _, r in totals.iterrows():
        apresentada = float(r["apresentada"])
        aprovada = float(r["aprovada"])
        taxa = round(100.0 * aprovada / apresentada, 1) if apresentada else 0.0
        totals_records.append(
            {
                "estabelecimento": r["estabelecimento_nome"],
                "apresentada": apresentada,
                "aprovada": aprovada,
                "taxa_aprovacao": taxa,
            }
        )
    totals_records.sort(key=lambda x: -x["apresentada"])

    # 3) Total geral por mes (visao consolidada da rede)
    total_mes = (
        df.groupby(["mes_atendimento", "conteudo"], as_index=False)["valor"].sum()
        .pivot_table(index="mes_atendimento", columns="conteudo", values="valor", fill_value=0)
        .reset_index()
    )
    for c in ("apresentada", "aprovada"):
        if c not in total_mes.columns:
            total_mes[c] = 0
    total_mes["mes_ord"] = total_mes["mes_atendimento"].apply(mes_chave)
    total_mes = total_mes.sort_values("mes_ord")
    total_mes_records = [
        {
            "mes": r["mes_atendimento"],
            "apresentada": float(r["apresentada"]),
            "aprovada": float(r["aprovada"]),
        }
        for _, r in total_mes.iterrows()
    ]

    # 4) Top CBOs por estabelecimento (top 8, resto agrupado em "Outros")
    cbo_tot = df.groupby(["estabelecimento_nome", "profissional_cbo", "conteudo"], as_index=False)["valor"].sum()
    cbo_apresentada = cbo_tot[cbo_tot["conteudo"] == "apresentada"]
    cbo_by_estab = {}
    for estab, grp in cbo_apresentada.groupby("estabelecimento_nome"):
        grp = grp.sort_values("valor", ascending=False)
        top = grp.head(8)
        outros_valor = grp["valor"].iloc[8:].sum() if len(grp) > 8 else 0.0
        itens = [{"cbo": row["profissional_cbo"], "valor": float(row["valor"])} for _, row in top.iterrows()]
        if outros_valor > 0:
            itens.append({"cbo": "Outros", "valor": float(outros_valor)})
        cbo_by_estab[estab] = itens

    # 5) Top CBOs geral (rede toda)
    cbo_rede = cbo_apresentada.groupby("profissional_cbo", as_index=False)["valor"].sum().sort_values("valor", ascending=False)
    top_rede = cbo_rede.head(10)
    outros_rede = cbo_rede["valor"].iloc[10:].sum() if len(cbo_rede) > 10 else 0.0
    cbo_rede_records = [{"cbo": r["profissional_cbo"], "valor": float(r["valor"])} for _, r in top_rede.iterrows()]
    if outros_rede > 0:
        cbo_rede_records.append({"cbo": "Outros", "valor": float(outros_rede)})

    payload = {
        "meses": list(meses),
        "estabelecimentos": list(estabelecimentos),
        "estabelecimentos_abrev": {e: abreviar_nome(e) for e in estabelecimentos},
        "mensal": monthly_records,
        "totais_estabelecimento": totals_records,
        "total_mes_rede": total_mes_records,
        "cbo_por_estabelecimento": cbo_by_estab,
        "cbo_rede": cbo_rede_records,
        "kpis": {
            "total_apresentada": float(df[df["conteudo"] == "apresentada"]["valor"].sum()),
            "total_aprovada": float(df[df["conteudo"] == "aprovada"]["valor"].sum()),
            "n_estabelecimentos": len(estabelecimentos),
            "n_meses": len(meses),
            "periodo_inicio": meses[0] if meses else None,
            "periodo_fim": meses[-1] if meses else None,
        },
    }
    total_apresentada = payload["kpis"]["total_apresentada"]
    total_aprovada = payload["kpis"]["total_aprovada"]
    payload["kpis"]["taxa_aprovacao_geral"] = round(100.0 * total_aprovada / total_apresentada, 1) if total_apresentada else 0.0

    os.makedirs(os.path.dirname(OUT_JS), exist_ok=True)
    with open(OUT_JS, "w", encoding="utf-8") as f:
        f.write("const TABNET_DATA = ")
        json.dump(payload, f, ensure_ascii=False, indent=None)
        f.write(";\n")

    print("JS de dados salvo em:", OUT_JS)
    print("Estabelecimentos:", len(estabelecimentos), "| Meses:", len(meses), "| Registros mensais:", len(monthly_records))


if __name__ == "__main__":
    main()
