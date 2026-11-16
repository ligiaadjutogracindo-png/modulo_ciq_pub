"""
Visão geral do CIQ (Dashboard): os cálculos, sem nada de tela.

Leitura em três níveis — laboratório → setor → teste — mas o cálculo continua por sistema analítico
(teste + equipamento + nível), como no resto do app. Um TESTE "precisa de ação" quando qualquer um dos
seus sistemas/níveis está fora da especificação no período de referência.

Período: Mensal ou Trimestral.
  - Mensal: cada sistema/nível usa o lote de maior N no mês (mesma regra das outras abas).
  - Trimestral: os meses do trimestre são agregados —
        CV pooled = raiz( Σ (Nᵢ − 1)·CVᵢ²  ÷  Σ (Nᵢ − 1) )
        Bias      = média dos Bias mensais ponderada pelo N
    e o status/Sigma/Erro Total do trimestre saem da MESMA função de avaliação usada no mês
    (processa, do app.py), com a especificação vigente no último mês do trimestre.

"Piorando, mas dentro da meta" (só para quem está dentro em TODAS as avaliações do período):
  - CV: subiu ≥ 10% em relação ao período anterior E o teste F (variâncias, 95%) é significativo.
        Sem o teste F, 10% de aumento aparece por acaso em ~1/3 dos sistemas todo mês.
  - Sigma: caiu ≥ 1 ponto E ficou abaixo de 6.
  Os limites (10%, 95%, 1 ponto) são ajustáveis na tela.
"""
from collections import defaultdict
import math

import pandas as pd
from scipy.stats import f as dist_f

MESES_NOME = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
              7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}
STATUS = ["Status CV", "Status Bias", "Status Erro Total", "Status Sigma"]
CHAVE_SISTEMA = ["Teste", "Equipamento", "NívelNum"]
FAIXAS_SIGMA = [("< 3", 0, 3), ("3 a 4", 3, 4), ("4 a 6", 4, 6), ("≥ 6", 6, math.inf)]


# ------------------------------------------------------------------ períodos
def chave_periodo(ano, mes, tipo):
    """Número sequencial do período (para ordenar e achar o anterior: anterior = chave − 1)."""
    return ano * 12 + (mes - 1) if tipo == "Mensal" else ano * 4 + (mes - 1) // 3


def rotulo_periodo(chave, tipo):
    if tipo == "Mensal":
        ano, mes0 = divmod(chave, 12)
        return f"{MESES_NOME[mes0 + 1]}/{str(ano)[2:]}"
    ano, t0 = divmod(chave, 4)
    return f"T{t0 + 1}/{str(ano)[2:]}"


# ------------------------------------------------------------------ agregação
def um_lote_por_mes(df):
    """Um registro por sistema/nível/mês: o lote de maior N (regra já usada nas abas CV, Bias...)."""
    base = df.dropna(subset=["NívelNum"]).copy()
    base["N"] = base["N"].fillna(0)
    idx = base.groupby(CHAVE_SISTEMA + ["_ordem_tempo"])["N"].idxmax()
    return base.loc[idx]


def cv_pooled(cvs, ns):
    """CV pooled ponderado pelos graus de liberdade (N − 1). Devolve (CV, graus de liberdade)."""
    pares = [(c, n - 1) for c, n in zip(cvs, ns) if pd.notna(c) and pd.notna(n) and n > 1]
    gl = sum(g for _, g in pares)
    if gl <= 0:
        return None, 0
    return math.sqrt(sum(g * c * c for c, g in pares) / gl), gl


def agrega(df, tipo, avalia, completo=True):
    """Uma linha por sistema (teste + equipamento + nível) e período, com CV, Bias, Erro Total, Sigma e
    status do período. `avalia(lista_de_registros)` é o processa do app (mesmas regras do mensal).
    No Mensal reaproveita a avaliação já feita; no Trimestral monta um registro agregado e reavalia."""
    base = um_lote_por_mes(df)
    base["_periodo"] = [chave_periodo(a, m, tipo) for a, m in zip(base["Ano"], base["Mês"])]
    linhas = []
    for (teste, equip, nivel, per), g in base.groupby(CHAVE_SISTEMA + ["_periodo"]):
        g = g.sort_values("_ordem_tempo")
        ult = g.iloc[-1]
        validos = g.dropna(subset=["CV (%)"])
        cv, gl = cv_pooled(validos["CV (%)"], validos["N"])
        if tipo == "Mensal":
            r = ult.to_dict()
        else:
            # registro agregado do trimestre, avaliado pela mesma função do mês
            com_bias = validos.dropna(subset=["Média", "Config. valor alvo"])
            com_bias = com_bias[com_bias["Config. valor alvo"] != 0]
            if com_bias.empty or cv is None:
                media = alvo = None
            else:
                pesos = com_bias["N"]
                bias = ((com_bias["Média"] - com_bias["Config. valor alvo"]) / com_bias["Config. valor alvo"])
                bias_w = (bias * pesos).sum() / pesos.sum()
                alvo = (com_bias["Config. valor alvo"] * pesos).sum() / pesos.sum()
                media = alvo * (1 + bias_w)
            sintetico = {
                "Arquivo": "", "Fonte CIQ": ult.get("Fonte CIQ"), "Módulo/Arquivo": ult["Módulo/Arquivo"],
                "Ano": int(ult["Ano"]), "Mês": int(ult["Mês"]), "Equipamento": equip, "Teste": teste,
                "Nível": ult["Nível"], "NívelNum": nivel, "Número de lote": "",
                # N total do período; só não pode cair na regra de N ≤ 5 se houver CV válido
                "N": float(validos["N"].sum()) if cv is not None else 0.0,
                "Config. valor alvo": alvo, "Média": media, "CV (%)": cv,
            }
            r = avalia([sintetico])[0]
        status = [r.get(s) for s in STATUS if completo or s != "Status Sigma"]
        avaliados = [s for s in status if s]
        linhas.append({
            "Período": per, "Teste": teste, "Equipamento": equip, "NívelNum": int(nivel),
            "Equipamento (nome)": ult["Equipamento (nome)"], "Setor": ult.get("Setor") or "Sem setor",
            "Meses": len(g), "N": validos["N"].sum(), "GL": gl,
            "CV (%)": r.get("CV (%)"), "CV Máximo": r.get("CV Máximo"),
            "Bias (%)": r.get("Bias (%)"), "Bias Observado": r.get("Bias Observado"),
            "Bias Máximo": r.get("Bias Máximo"),
            "Erro Total": r.get("Erro Total Observado"), "ETM": r.get("ETM (para comparação)"),
            "Sigma": r.get("Sigma Mensal") if completo else None,
            **{s: r.get(s) for s in STATUS},
            "Situação": (None if not avaliados else "Fora" if "VERMELHO" in avaliados else "Dentro"),
        })
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------ "piorando"
def p_teste_f(cv_atual, gl_atual, cv_ant, gl_ant):
    """Teste F unilateral: a variância (CV²) do período atual é maior que a do anterior?"""
    if not cv_atual or not cv_ant or gl_atual < 1 or gl_ant < 1:
        return None
    return float(dist_f.sf((cv_atual / cv_ant) ** 2, gl_atual, gl_ant))


def marca_piorando(agr, lim_cv=0.10, confianca=0.95, queda_sigma=1.0, completo=True):
    """Acrescenta ao agregado: CV/Sigma do período anterior, p do teste F e o motivo de "piorando"
    (só para sistemas com Situação = Dentro)."""
    if agr.empty:
        return agr.assign(**{c: None for c in ("CV anterior", "Sigma anterior", "p (teste F)", "Piorando")})
    anterior = {(r["Teste"], r["Equipamento"], r["NívelNum"], r["Período"]): r for r in agr.to_dict("records")}
    motivos, cv_ant_l, sig_ant_l, p_l = [], [], [], []
    for r in agr.to_dict("records"):
        a = anterior.get((r["Teste"], r["Equipamento"], r["NívelNum"], r["Período"] - 1))
        cv, sigma = r["CV (%)"], r["Sigma"]
        cv_a = a["CV (%)"] if a else None
        sig_a = a["Sigma"] if a else None
        p = p_teste_f(cv, r["GL"], cv_a, a["GL"] if a else 0)
        cv_ant_l.append(cv_a), sig_ant_l.append(sig_a), p_l.append(p)
        m = []
        if r["Situação"] == "Dentro" and a:
            if cv and cv_a and cv >= cv_a * (1 + lim_cv) and p is not None and p < 1 - confianca:
                m.append(f"CV {cv_a:.2f}% → {cv:.2f}%")
            if (completo and pd.notna(sigma) and pd.notna(sig_a)
                    and sig_a - sigma >= queda_sigma and sigma < 6):
                m.append(f"Sigma {sig_a:.1f} → {sigma:.1f}")
        motivos.append("; ".join(m) or None)
    return agr.assign(**{"CV anterior": cv_ant_l, "Sigma anterior": sig_ant_l, "p (teste F)": p_l,
                         "Piorando": motivos})


# ------------------------------------------------------------------ leituras para a tela
def uso_da_meta(row):
    """Quanto da meta o sistema usa (0–1+): o maior entre CV/CV Máx, Bias/Bias Máx e ET/ETM."""
    fr = [abs(v) / lim for v, lim in ((row["CV (%)"], row["CV Máximo"]),
                                      (row["Bias Observado"], row["Bias Máximo"]),
                                      (row["Erro Total"], row["ETM"]))
          if pd.notna(v) and pd.notna(lim) and lim]
    return max(fr) if fr else None


def resumo_testes(agr, periodo):
    """Uma linha por teste no período: situação (pior entre sistemas), quantos sistemas fora, pior Sigma."""
    p = agr[(agr["Período"] == periodo) & agr["Situação"].notna()]
    if p.empty:
        return pd.DataFrame(columns=["Teste", "Setor", "Situação", "Sistemas", "Sistemas fora", "Pior Sigma",
                                     "Piorando"])
    return (p.groupby(["Teste", "Setor"])
            .agg(Sistemas=("Situação", "size"),
                 **{"Sistemas fora": ("Situação", lambda s: int((s == "Fora").sum()))},
                 **{"Pior Sigma": ("Sigma", "min")},
                 Piorando=("Piorando", lambda s: s.notna().any()))
            .reset_index()
            .assign(Situação=lambda d: d["Sistemas fora"].map(lambda n: "Fora" if n else "Dentro")))


def evolucao_setor(agr):
    """% dos sistemas avaliados dentro da especificação, por setor e período (e 'Laboratório' = todos)."""
    av = agr[agr["Situação"].notna()]
    if av.empty:
        return pd.DataFrame(columns=["Setor", "Período", "% dentro", "Avaliados"])
    todos = av.assign(Setor="Laboratório")
    out = (pd.concat([av, todos]).groupby(["Setor", "Período"])
           .agg(Avaliados=("Situação", "size"), Dentro=("Situação", lambda s: (s == "Dentro").sum()))
           .reset_index())
    out["% dentro"] = out["Dentro"] / out["Avaliados"] * 100
    return out


def faixas_sigma(resumo):
    """Quantos testes caem em cada faixa do pior Sigma."""
    s = resumo["Pior Sigma"].dropna()
    return {nome: int(((s >= lo) & (s < hi)).sum()) for nome, lo, hi in FAIXAS_SIGMA}


def serie_cv(base_mensal, teste, equip, nivel, ate_ordem, meses=12):
    """CV mês a mês de um sistema (últimos `meses` até o mês de referência) — para a minilinha."""
    s = base_mensal[(base_mensal["Teste"] == teste) & (base_mensal["Equipamento"] == equip)
                    & (base_mensal["NívelNum"] == nivel) & (base_mensal["_ordem_tempo"] <= ate_ordem)]
    s = s.sort_values("_ordem_tempo").tail(meses)
    return [round(v, 2) for v in s["CV (%)"] if pd.notna(v)]
