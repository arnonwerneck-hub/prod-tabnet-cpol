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


PSEUDONIMOS = {
    "SMS CENTRO CARIOCA DE REABILITACAO DA ZONA OESTE AP 52": "CCRZO",
    "SMS CENTRO CARIOCA DE ESPECIALIDADE DA ZONA OESTE AP 52": "CCEZO",
    "SMS POLICLINICA MANOEL GUILHERME PAM BANGU AP 51": "PMGSF",
    "SMS POLICLINICA NEWTON ALVES CARDOZO AP 31": "PNAC",
    "SMS POLICLINICA LINCOLN DE FREITAS FILHO AP 53": "PLFF",
    "SMS POLICLINICA NEWTON BETHLEM AP 40": "PNB",
    "SMS POLICLINICA RODOLPHO ROCCO AP 32": "PRR",
    "SMS POLICLINICA CARLOS ALBERTO NASCIMENTO AP 52": "PCAN",
    "SMS POLICLINICA JOSE PARANHOS FONTENELLE AP 31": "PJPF",
    "SMS POLICLINICA ANTONIO RIBEIRO NETTO AP 10": "PARN",
    "SMS POLICLINICA HELIO PELLEGRINO AP 22": "PHP",
    "SMS POLICLINICA ROCHA MAIA AP 21": "PRM",
}


def abreviar_nome(nome):
    if nome in PSEUDONIMOS:
        return PSEUDONIMOS[nome]
    nome = nome.replace("SMS POLICLINICA ", "").replace("SMS CENTRO CARIOCA DE ", "C.C. ")
    return nome.strip()


def eh_medico(profissional_cbo):
    """Classifica pelo nome da ocupacao (apos o codigo CBO): 'Medico...' = medico.
    Ex.: '225135 Medico dermatologista' -> medico; '221205 Biomedico' -> nao medico."""
    partes = profissional_cbo.split(" ", 1)
    nome_ocupacao = (partes[1] if len(partes) > 1 else profissional_cbo).strip().lower()
    return nome_ocupacao.startswith("médico") or nome_ocupacao.startswith("medico")


def main():
    df = pd.read_csv(IN_CSV, encoding="utf-8")
    df = df[df["mes_atendimento"].apply(lambda x: mes_chave(x)[0] >= 2026)].copy()
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
        taxa = round(100.0 * aprovada / apresentada, 2) if apresentada else 0.0
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
        taxa = round(100.0 * aprovada / apresentada, 2) if apresentada else 0.0
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

    # 4) Top CBOs por estabelecimento, separados em Medicos / Nao-Medicos (top 8 cada, resto em "Outros")
    cbo_tot = df.groupby(["estabelecimento_nome", "profissional_cbo", "conteudo"], as_index=False)["valor"].sum()
    cbo_apresentada = cbo_tot[cbo_tot["conteudo"] == "apresentada"].copy()
    cbo_apresentada["categoria"] = cbo_apresentada["profissional_cbo"].apply(lambda c: "medicos" if eh_medico(c) else "nao_medicos")

    def montar_top(grp, n=8):
        grp = grp.sort_values("valor", ascending=False)
        top = grp.head(n)
        outros_valor = grp["valor"].iloc[n:].sum() if len(grp) > n else 0.0
        itens = [{"cbo": row["profissional_cbo"], "valor": float(row["valor"])} for _, row in top.iterrows()]
        if outros_valor > 0:
            itens.append({"cbo": "Outros", "valor": float(outros_valor)})
        return itens

    cbo_by_estab = {"medicos": {}, "nao_medicos": {}}
    for (estab, categoria), grp in cbo_apresentada.groupby(["estabelecimento_nome", "categoria"]):
        cbo_by_estab[categoria][estab] = montar_top(grp)

    # 5) Top CBOs geral (rede toda), separados em Medicos / Nao-Medicos
    cbo_rede_por_categoria = {}
    for categoria, grp in cbo_apresentada.groupby("categoria"):
        agregado = grp.groupby("profissional_cbo", as_index=False)["valor"].sum()
        cbo_rede_por_categoria[categoria] = montar_top(agregado, n=10)

    payload = {
        "meses": list(meses),
        "estabelecimentos": list(estabelecimentos),
        "estabelecimentos_abrev": {e: abreviar_nome(e) for e in estabelecimentos},
        "mensal": monthly_records,
        "totais_estabelecimento": totals_records,
        "total_mes_rede": total_mes_records,
        "cbo_por_estabelecimento_medicos": cbo_by_estab["medicos"],
        "cbo_por_estabelecimento_nao_medicos": cbo_by_estab["nao_medicos"],
        "cbo_rede_medicos": cbo_rede_por_categoria.get("medicos", []),
        "cbo_rede_nao_medicos": cbo_rede_por_categoria.get("nao_medicos", []),
        "kpis": {
            "total_apresentada": float(df[df["conteudo"] == "apresentada"]["valor"].sum()),
            "total_aprovada": float(df[df["conteudo"] == "aprovada"]["valor"].sum()),
            "n_estabelecimentos": len(estabelecimentos),
            "n_meses": len(meses),
            "n_cbos": int(df[df["conteudo"] == "apresentada"]["profissional_cbo"].nunique()),
            "periodo_inicio": meses[0] if meses else None,
            "periodo_fim": meses[-1] if meses else None,
        },
    }
    total_apresentada = payload["kpis"]["total_apresentada"]
    total_aprovada = payload["kpis"]["total_aprovada"]
    n_meses = payload["kpis"]["n_meses"] or 1
    payload["kpis"]["taxa_aprovacao_geral"] = round(100.0 * total_aprovada / total_apresentada, 1) if total_apresentada else 0.0
    payload["kpis"]["media_mensal_apresentada"] = round(total_apresentada / n_meses, 1)
    payload["kpis"]["media_mensal_aprovada"] = round(total_aprovada / n_meses, 1)

    os.makedirs(os.path.dirname(OUT_JS), exist_ok=True)
    with open(OUT_JS, "w", encoding="utf-8") as f:
        f.write("const TABNET_DATA = ")
        json.dump(payload, f, ensure_ascii=False, indent=None)
        f.write(";\n")

    print("JS de dados salvo em:", OUT_JS)
    print("Estabelecimentos:", len(estabelecimentos), "| Meses:", len(meses), "| Registros mensais:", len(monthly_records))


if __name__ == "__main__":
    main()
